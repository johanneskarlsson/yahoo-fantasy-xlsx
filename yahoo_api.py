import os
import xmltodict
from requests_oauthlib import OAuth2Session
from dotenv import load_dotenv
import json
import logging
import webbrowser
from typing import Union, List, Dict, Any

load_dotenv()

class YahooFantasyAPI:
    def __init__(self):
        # Credentials / league identifiers
        self.client_id = os.getenv('YAHOO_CLIENT_ID')
        self.client_secret = os.getenv('YAHOO_CLIENT_SECRET')
        self.league_id = os.getenv('LEAGUE_ID')

        # OAuth endpoints
        self.authorization_base_url = 'https://api.login.yahoo.com/oauth2/request_auth'
        self.token_url = 'https://api.login.yahoo.com/oauth2/get_token'
        self.redirect_uri = 'https://developers.google.com/oauthplayground'

        # Session state
        self.session: OAuth2Session | None = None
        self.year_id: str | None = None

        # Logging
        self.logger = logging.getLogger(__name__)

        # Caches
        self._player_name_cache: dict[str, str] = {}

    def _extract_dict_value(self, data: Union[Dict, Any], key: str = None) -> str:
        """Extract value from dictionary structure that may have #text key"""
        if isinstance(data, dict):
            if key and key in data:
                return self._extract_dict_value(data[key])
            return data.get('#text', data)
        return str(data) if data is not None else ''

    def _ensure_list(self, data: Union[List, Dict, None]) -> List:
        """Ensure data is returned as a list"""
        if data is None:
            return []
        if isinstance(data, dict):
            return [data]
        return data if isinstance(data, list) else []

    def _make_api_request(self, url: str) -> Dict[str, Any]:
        """Make authenticated API request with common error handling"""
        if not self.session:
            self.ensure_authenticated()

        try:
            response = self.session.get(url)
            response.raise_for_status()
            return xmltodict.parse(response.text)
        except Exception as e:
            msg = str(e)
            if 'token_expired' in msg.lower():
                self.logger.warning(f"Token appears expired for {url}; will trigger refresh if possible.")
            else:
                self.logger.error(f"API request failed for {url}: {e}")
            raise

    def authenticate(self):
        """Handle OAuth2 authentication flow"""
        yahoo = OAuth2Session(
            self.client_id,
            redirect_uri=self.redirect_uri,
            scope=['fspt-r']
        )

        # Get authorization URL
        authorization_url, _ = yahoo.authorization_url(
            self.authorization_base_url,
            access_type="offline",
            prompt="select_account"
        )

        print(f'\nPlease go to {authorization_url} and authorize access.')
        print('After authorizing, you will be redirected to a page that may show an error.')
        print('Copy the authorization code from the URL and paste it below.')

        webbrowser.open(authorization_url)

        code = input('\nEnter the authorization code: ').strip()

        # Fetch token
        token = yahoo.fetch_token(
            self.token_url,
            code=code,
            client_secret=self.client_secret
        )

        self.session = yahoo

        # Save token to file for future use
        with open('token.json', 'w') as f:
            json.dump(token, f)

        return True

    def load_token(self):
        """Load saved token if exists"""
        try:
            with open('token.json', 'r') as f:
                token = json.load(f)

            self.session = OAuth2Session(
                self.client_id,
                token=token
            )
            return True
        except FileNotFoundError:
            return False

    def refresh_token_if_needed(self):
        """Refresh token if it's expired"""
        if not self.session:
            return False

        try:
            # Try a simple API call to check if token is valid
            self.get_game_key()
            return True
        except Exception:
            # Token might be expired, try to refresh
            try:
                token = self.session.refresh_token(
                    self.token_url,
                    client_id=self.client_id,
                    client_secret=self.client_secret
                )

                # Save refreshed token
                with open('token.json', 'w') as f:
                    json.dump(token, f)

                return True
            except Exception as e:
                self.logger.error(f"Failed to refresh token: {e}")
                return False

    def ensure_authenticated(self):
        """Ensure we have a valid authenticated session"""
        if not self.load_token():
            self.authenticate()
        elif not self.refresh_token_if_needed():
            self.authenticate()

    def get_game_key(self):
        """Get the current year's game key"""
        url = "https://fantasysports.yahooapis.com/fantasy/v2/game/nhl"

        data = self._make_api_request(url)
        game_data = data['fantasy_content']['game']
        game_key = self._extract_dict_value(game_data, 'game_key')

        self.year_id = game_key
        self.logger.debug(f"Game key: {game_key}")
        return game_key

    def get_draft_results(self):
        """Get draft results from Yahoo API"""
        if not self.year_id:
            self.get_game_key()

        url = f"https://fantasysports.yahooapis.com/fantasy/v2/league/{self.year_id}.l.{self.league_id}/draftresults"

        try:
            data = self._make_api_request(url)
            league_data = data['fantasy_content']['league']

            self.logger.debug(f"Draft results response structure: {list(league_data.keys())}")

            if 'draft_results' not in league_data:
                self.logger.debug("No draft_results found in response")
                return []

            draft_results_data = league_data['draft_results']

            if draft_results_data is None or 'draft_result' not in draft_results_data:
                self.logger.debug("No draft_result found in draft_results")
                return []

            return self._ensure_list(draft_results_data['draft_result'])

        except Exception as e:
            self.logger.error(f"Error getting draft results: {e}")
            return []

    def get_league_settings(self):
        """Get league settings from Yahoo API"""
        if not self.year_id:
            self.get_game_key()

        url = f"https://fantasysports.yahooapis.com/fantasy/v2/league/{self.year_id}.l.{self.league_id}/settings"

        try:
            data = self._make_api_request(url)
            league_data = data['fantasy_content']['league']

            if 'settings' not in league_data:
                self.logger.debug("No settings found in response")
                return {}

            settings = league_data['settings']
            self.logger.debug(f"Raw settings structure keys: {list(settings.keys())}")

            # Extract key league settings with safe extraction
            league_settings = {
                'league_name': self._extract_dict_value(league_data, 'name'),
                'league_type': self._extract_dict_value(settings, 'draft_type'),
                'scoring_type': self._extract_dict_value(settings, 'scoring_type'),
                'max_teams': self._extract_dict_value(settings, 'max_teams'),
                'num_playoff_teams': self._extract_dict_value(settings, 'num_playoff_teams'),
                'playoff_start_week': self._extract_dict_value(settings, 'playoff_start_week'),
                'waiver_type': self._extract_dict_value(settings, 'waiver_type'),
                'trade_end_date': self._extract_dict_value(settings, 'trade_end_date'),
                'roster_positions': [],
                'stat_categories': []
            }

            # Extract roster positions safely
            try:
                if 'roster_positions' in settings:
                    roster_data = settings['roster_positions']
                    if isinstance(roster_data, dict) and 'roster_position' in roster_data:
                        positions = roster_data['roster_position']
                        if isinstance(positions, list):
                            for pos in positions:
                                if isinstance(pos, dict):
                                    league_settings['roster_positions'].append({
                                        'position': pos.get('position', ''),
                                        'count': pos.get('count', '')
                                    })
                        elif isinstance(positions, dict):
                            league_settings['roster_positions'].append({
                                'position': positions.get('position', ''),
                                'count': positions.get('count', '')
                            })
            except Exception as e:
                self.logger.warning(f"Error extracting roster positions: {e}")

            # Extract stat categories safely
            try:
                if 'stat_categories' in settings:
                    stat_data = settings['stat_categories']
                    if isinstance(stat_data, dict) and 'stats' in stat_data:
                        stats_section = stat_data['stats']
                        if isinstance(stats_section, dict) and 'stat' in stats_section:
                            stats = stats_section['stat']
                            if isinstance(stats, list):
                                for stat in stats:
                                    if isinstance(stat, dict):
                                        league_settings['stat_categories'].append({
                                            'stat_id': stat.get('stat_id', ''),
                                            'name': stat.get('name', ''),
                                            'display_name': stat.get('display_name', ''),
                                            'position_type': stat.get('position_type', ''),
                                            'value': self._get_stat_modifier_value(settings, stat.get('stat_id', ''))
                                        })
                            elif isinstance(stats, dict):
                                league_settings['stat_categories'].append({
                                    'stat_id': stats.get('stat_id', ''),
                                    'name': stats.get('name', ''),
                                    'display_name': stats.get('display_name', ''),
                                    'position_type': stats.get('position_type', ''),
                                    'value': self._get_stat_modifier_value(settings, stats.get('stat_id', ''))
                                })
            except Exception as e:
                self.logger.warning(f"Error extracting stat categories: {e}")

            self.logger.debug(f"Retrieved league settings for: {league_settings['league_name']}")
            return league_settings

        except Exception as e:
            self.logger.error(f"Error getting league settings: {e}")
            return {}

    def _get_stat_modifier_value(self, settings, stat_id):
        """Get stat modifier value for a given stat ID"""
        try:
            if 'stat_modifiers' in settings:
                modifiers = settings['stat_modifiers']
                if isinstance(modifiers, dict) and 'stats' in modifiers:
                    stats_section = modifiers['stats']
                    if isinstance(stats_section, dict) and 'stat' in stats_section:
                        stats = stats_section['stat']
                        if isinstance(stats, list):
                            for stat in stats:
                                if isinstance(stat, dict) and stat.get('stat_id') == stat_id:
                                    return stat.get('value', '')
                        elif isinstance(stats, dict) and stats.get('stat_id') == stat_id:
                            return stats.get('value', '')
        except Exception:
            pass
        return ''

    def get_teams_data(self):
        """Get teams data from Yahoo API"""
        if not self.year_id:
            self.get_game_key()

        url = f"https://fantasysports.yahooapis.com/fantasy/v2/league/{self.year_id}.l.{self.league_id}/teams"

        try:
            data = self._make_api_request(url)
            teams = self._ensure_list(data['fantasy_content']['league']['teams']['team'])

            teams_data = []
            for team in teams:
                team_key = self._extract_dict_value(team, 'team_key')
                team_id = self._extract_dict_value(team, 'team_id')
                team_name = self._extract_dict_value(team, 'name')
                manager_name = self._extract_dict_value(team['managers']['manager'], 'nickname')

                teams_data.append([team_key, team_id, team_name, manager_name])

            return teams_data

        except Exception as e:
            self.logger.error(f"Error getting teams data: {e}")
            return []


    def get_player_draft_analysis(self):
        """Get draft analysis data including ADP if available - fetch in batches"""
        if not self.year_id:
            self.get_game_key()

        # Base endpoints based on the correct Yahoo API structure
        base_endpoints = [
            f"https://fantasysports.yahooapis.com/fantasy/v2/league/{self.year_id}.l.{self.league_id}/players;position=ALL;sort=average_pick;out=auction_values,ranks;ranks=season;ranks_by_position=season/draft_analysis;cut_types=diamond;slices=last7days"
        ]

        for base_url in base_endpoints:
            try:
                self.logger.debug(f"Trying draft analysis with batching from: {base_url}")

                all_draft_data = []
                start = 0
                batch_size = 25  # Yahoo API seems to limit to 25
                max_iterations = 40  # 40 * 25 = 1000 players max

                for iteration in range(max_iterations):
                    # Insert the start and count parameters into the URL
                    url = base_url.replace(";position=ALL;", f";position=ALL;start={start};count={batch_size};")

                    try:
                        data = self._make_api_request(url)

                        # Navigate the response structure
                        if 'fantasy_content' in data:
                            content = data['fantasy_content']
                            players_data = None

                            if 'league' in content and 'players' in content['league']:
                                players_data = content['league']['players']
                            elif 'game' in content and 'players' in content['game']:
                                players_data = content['game']['players']

                            if players_data and 'player' in players_data:
                                players = self._ensure_list(players_data['player'])

                                if not players:
                                    self.logger.debug(f"No more players found at start={start}. Stopping.")
                                    break

                                batch_draft_data = []
                                for player in players:
                                    player_info = self._extract_player_draft_info(player)
                                    if player_info:
                                        batch_draft_data.append(player_info)

                                all_draft_data.extend(batch_draft_data)
                                self.logger.debug(f"Batch {iteration + 1}: Got {len(batch_draft_data)} players, total: {len(all_draft_data)}")

                                # If we got fewer players than requested, we've reached the end
                                if len(players) < batch_size:
                                    self.logger.debug(f"Reached end of data. Got {len(players)} in final batch.")
                                    break
                            else:
                                self.logger.debug(f"No players data found at start={start}")
                                break

                    except Exception as e:
                        self.logger.warning(f"Error in batch {iteration + 1} at start={start}: {e}")
                        break

                    start += batch_size

                if all_draft_data:
                    self.logger.debug(f"Successfully extracted draft data for {len(all_draft_data)} players from {base_url}")
                    return all_draft_data

            except Exception as e:
                self.logger.debug(f"Failed to get batched data from {base_url}: {e}")
                continue

        self.logger.debug("No draft analysis data available from any endpoint")
        return []

    # -------------------- Lightweight Player Lookup --------------------
    def get_player_name(self, player_key: str) -> str:
        """Fetch a single player's full name by player_key with simple caching.

        Uses the player endpoint: /fantasy/v2/player/{player_key}
        """
        if not player_key:
            return ""
        if player_key in self._player_name_cache:
            return self._player_name_cache[player_key]
        if not self.year_id:
            self.get_game_key()
        url = f"https://fantasysports.yahooapis.com/fantasy/v2/player/{player_key}"
        try:
            data = self._make_api_request(url)
            # Navigate to player -> name -> full
            player = data.get('fantasy_content', {}).get('player')
            if not player:
                # Sometimes structure is fantasy_content -> players -> player
                players = data.get('fantasy_content', {}).get('players', {})
                player = players.get('player') if isinstance(players, dict) else None
            full_name = ""
            if player and isinstance(player, dict) and 'name' in player:
                name_dict = player['name']
                if isinstance(name_dict, dict):
                    full_name = name_dict.get('full') or name_dict.get('first') or ""
            full_name = full_name or "(unknown)"
            self._player_name_cache[player_key] = full_name
            return full_name
        except Exception as e:
            self.logger.warning(f"Failed to fetch player name for {player_key}: {e}")
            return ""

    def _extract_player_draft_info(self, player):
        """Extract draft-related information from a player object"""
        try:
            player_key = self._extract_dict_value(player, 'player_key')
            full_name = self._extract_dict_value(player['name'], 'full')

            # Extract draft analysis data
            draft_info = {}
            if 'draft_analysis' in player:
                draft_analysis = player['draft_analysis']
                draft_info['average_pick'] = self._extract_dict_value(draft_analysis, 'average_pick')
                draft_info['average_round'] = self._extract_dict_value(draft_analysis, 'average_round')
                draft_info['percent_drafted'] = self._extract_dict_value(draft_analysis, 'percent_drafted')
                draft_info['preseason_average_pick'] = self._extract_dict_value(draft_analysis, 'preseason_average_pick')
                draft_info['preseason_percent_drafted'] = self._extract_dict_value(draft_analysis, 'preseason_percent_drafted')

            # Extract auction values
            draft_info['projected_auction_value'] = self._extract_dict_value(player, 'projected_auction_value')
            draft_info['average_auction_cost'] = self._extract_dict_value(player, 'average_auction_cost')

            # Extract rankings
            season_rank = ''
            position_rank = ''
            if 'player_ranks' in player and 'player_rank' in player['player_ranks']:
                ranks = self._ensure_list(player['player_ranks']['player_rank'])
                for rank in ranks:
                    rank_type = rank.get('rank_type', '')
                    rank_season = rank.get('rank_season', '')
                    rank_position = rank.get('rank_position', '')

                    # Get current season overall rank
                    if rank_type == 'S' and rank_season == '2025' and not rank_position:
                        season_rank = self._extract_dict_value(rank, 'rank_value')
                    # Get current season position rank
                    elif rank_type == 'S' and rank_season == '2025' and rank_position:
                        position_rank = self._extract_dict_value(rank, 'rank_value')

            return [
                player_key,
                full_name,
                self._extract_dict_value(player, 'editorial_team_abbr'),
                self._extract_dict_value(player, 'display_position'),
                draft_info.get('average_pick', ''),
                draft_info.get('average_round', ''),
                draft_info.get('percent_drafted', ''),
                draft_info.get('projected_auction_value', ''),
                draft_info.get('average_auction_cost', ''),
                season_rank,
                position_rank,
                draft_info.get('preseason_average_pick', ''),
                draft_info.get('preseason_percent_drafted', '')
            ]

        except Exception as e:
            self.logger.error(f"Error extracting draft info for player: {e}")
            return None
