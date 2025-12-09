"""Vercel Serverless Function for transaction handling."""

from http.server import BaseHTTPRequestHandler
import json
import os
import base64
import urllib.request
from urllib.error import HTTPError
from datetime import datetime
import uuid

# GitHub repo info
GITHUB_OWNER = os.environ.get("REPO_OWNER") or os.environ.get("GITHUB_OWNER", "gansel51")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "opfl-autoscorer")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")

TRADE_DEADLINE_WEEK = 12
MAX_TAXI_SQUAD = 3


def get_team_password(team_abbrev: str) -> str | None:
    """Get the password for a team from environment variables."""
    env_key = f"TEAM_PASSWORD_{team_abbrev.replace('/', '_')}"
    return os.environ.get(env_key)


def github_api_request(path: str, method: str = "GET", data: dict = None) -> tuple[bool, dict | str]:
    """Make a GitHub API request."""
    github_token = os.environ.get("SKYNET_PAT") or os.environ.get("GITHUB_TOKEN")
    if not github_token:
        return False, "Server configuration error - no GitHub token"
    
    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{path}"
    
    headers = {
        "Authorization": f"Bearer {github_token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
        "User-Agent": "OPFL-Transaction-Bot"
    }
    
    try:
        if method == "GET":
            req = urllib.request.Request(api_url, headers=headers)
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode())
                content = json.loads(base64.b64decode(result["content"]).decode())
                return True, {"content": content, "sha": result["sha"]}
        
        elif method == "PUT":
            req = urllib.request.Request(api_url, headers=headers)
            current_sha = None
            try:
                with urllib.request.urlopen(req) as response:
                    result = json.loads(response.read().decode())
                    current_sha = result["sha"]
            except HTTPError as e:
                if e.code != 404:
                    return False, f"Failed to get file: {e}"
            
            update_data = {
                "message": data.get("message", "Update file"),
                "content": base64.b64encode(json.dumps(data["content"], indent=2).encode()).decode(),
                "branch": GITHUB_BRANCH
            }
            if current_sha:
                update_data["sha"] = current_sha
            
            req = urllib.request.Request(
                api_url,
                data=json.dumps(update_data).encode(),
                headers=headers,
                method="PUT"
            )
            with urllib.request.urlopen(req) as response:
                return True, "File updated successfully"
                
    except HTTPError as e:
        error_body = e.read().decode() if hasattr(e, 'read') else str(e)
        return False, f"GitHub API error: {error_body}"
    except Exception as e:
        return False, str(e)
    
    return False, "Unknown error"


def validate_team(team: str, password: str) -> tuple[bool, str]:
    """Validate team password."""
    if not team or not password:
        return False, "Missing team or password"
    
    expected = get_team_password(team)
    if not expected:
        return False, "Team not configured"
    
    if password != expected:
        return False, "Invalid password"
    
    return True, "Valid"


def handle_taxi_activation(data: dict) -> tuple[int, dict]:
    """Handle taxi squad activation (promote player from taxi to active roster)."""
    team = data.get("team")
    password = data.get("password")
    player_to_activate = data.get("player_to_activate")
    player_to_release = data.get("player_to_release")
    week = data.get("week")
    
    valid, msg = validate_team(team, password)
    if not valid:
        return 401, {"error": msg}
    
    if not all([player_to_activate, player_to_release, week]):
        return 400, {"error": "Missing required fields"}
    
    # Get current roster
    success, result = github_api_request("data/rosters.json")
    if not success:
        return 500, {"error": result}
    
    rosters = result["content"]
    roster, taxi = get_roster_and_taxi(rosters, team)
    
    # Validate player_to_activate is in taxi
    taxi_player = next((p for p in taxi if p["name"] == player_to_activate), None)
    if not taxi_player:
        return 400, {"error": f"{player_to_activate} is not on your taxi squad"}
    
    # Validate player_to_release is in roster and same position
    roster_player = next((p for p in roster if p["name"] == player_to_release), None)
    if not roster_player:
        return 400, {"error": f"{player_to_release} is not on your active roster"}
    
    if taxi_player["position"] != roster_player["position"]:
        return 400, {"error": f"Position mismatch: {taxi_player['position']} vs {roster_player['position']}"}
    
    # Execute the swap
    taxi = [p for p in taxi if p["name"] != player_to_activate]
    roster = [p for p in roster if p["name"] != player_to_release]
    roster.append(taxi_player)
    
    set_roster_and_taxi(rosters, team, roster, taxi)
    
    # Save updated rosters
    success, msg = github_api_request("data/rosters.json", "PUT", {
        "message": f"Taxi activation: {team} activates {player_to_activate}, releases {player_to_release}",
        "content": rosters
    })
    
    if not success:
        return 500, {"error": msg}
    
    # Add to transaction log
    add_transaction_log({
        "type": "taxi_activation",
        "team": team,
        "activated": player_to_activate,
        "released": player_to_release,
        "week": week,
        "timestamp": datetime.utcnow().isoformat()
    })
    
    return 200, {"success": True, "message": f"Activated {player_to_activate}, released {player_to_release}"}


def handle_propose_trade(data: dict) -> tuple[int, dict]:
    """Handle trade proposal."""
    team = data.get("team")
    password = data.get("password")
    trade_partner = data.get("trade_partner")
    give_players = data.get("give_players", [])
    give_picks = data.get("give_picks", [])
    receive_players = data.get("receive_players", [])
    receive_picks = data.get("receive_picks", [])
    current_week = data.get("current_week", 1)
    
    valid, msg = validate_team(team, password)
    if not valid:
        return 401, {"error": msg}
    
    if not trade_partner:
        return 400, {"error": "Must specify trade partner"}
    
    if not (give_players or give_picks) and not (receive_players or receive_picks):
        return 400, {"error": "Trade must include players or picks"}
    
    if current_week >= TRADE_DEADLINE_WEEK:
        return 400, {"error": f"Trade deadline has passed (Week {TRADE_DEADLINE_WEEK})"}
    
    success, result = github_api_request("data/pending_trades.json")
    if not success:
        pending = {"trades": [], "trade_deadline_week": TRADE_DEADLINE_WEEK}
    else:
        pending = result["content"]
    
    trade = {
        "id": str(uuid.uuid4())[:8],
        "proposer": team,
        "partner": trade_partner,
        "proposer_gives": {"players": give_players, "picks": give_picks},
        "proposer_receives": {"players": receive_players, "picks": receive_picks},
        "status": "pending",
        "proposed_at": datetime.utcnow().isoformat(),
        "week": current_week
    }
    
    pending["trades"].append(trade)
    
    success, msg = github_api_request("data/pending_trades.json", "PUT", {
        "message": f"Trade proposed: {team} to {trade_partner}",
        "content": pending
    })
    
    if not success:
        return 500, {"error": msg}
    
    return 200, {"success": True, "message": f"Trade proposed to {trade_partner}", "trade_id": trade["id"]}


def get_roster_and_taxi(rosters: dict, team: str) -> tuple[list, list]:
    """Get roster and taxi squad from rosters data, handling both formats."""
    team_data = rosters.get(team, [])
    if isinstance(team_data, list):
        return team_data, []
    else:
        return team_data.get("roster", []), team_data.get("taxi_squad", [])


def set_roster_and_taxi(rosters: dict, team: str, roster: list, taxi: list):
    """Set roster and taxi squad, preserving the original format."""
    if team in rosters and isinstance(rosters[team], dict):
        rosters[team] = {"roster": roster, "taxi_squad": taxi}
    else:
        rosters[team] = roster


def execute_trade(trade: dict) -> tuple[bool, str]:
    """Execute a trade by swapping players between teams."""
    proposer = trade["proposer"]
    partner = trade["partner"]
    proposer_gives = trade["proposer_gives"]
    proposer_receives = trade["proposer_receives"]
    
    success, result = github_api_request("data/rosters.json")
    if not success:
        return False, f"Failed to get rosters: {result}"
    
    rosters = result["content"]
    
    proposer_roster, proposer_taxi = get_roster_and_taxi(rosters, proposer)
    partner_roster, partner_taxi = get_roster_and_taxi(rosters, partner)
    
    players_to_partner = []
    for player_name in proposer_gives.get("players", []):
        player = next((p for p in proposer_roster if p["name"] == player_name), None)
        if player:
            proposer_roster = [p for p in proposer_roster if p["name"] != player_name]
            players_to_partner.append(player)
        else:
            player = next((p for p in proposer_taxi if p["name"] == player_name), None)
            if player:
                proposer_taxi = [p for p in proposer_taxi if p["name"] != player_name]
                players_to_partner.append(player)
    
    players_to_proposer = []
    for player_name in proposer_receives.get("players", []):
        player = next((p for p in partner_roster if p["name"] == player_name), None)
        if player:
            partner_roster = [p for p in partner_roster if p["name"] != player_name]
            players_to_proposer.append(player)
        else:
            player = next((p for p in partner_taxi if p["name"] == player_name), None)
            if player:
                partner_taxi = [p for p in partner_taxi if p["name"] != player_name]
                players_to_proposer.append(player)
    
    partner_roster.extend(players_to_partner)
    proposer_roster.extend(players_to_proposer)
    
    set_roster_and_taxi(rosters, proposer, proposer_roster, proposer_taxi)
    set_roster_and_taxi(rosters, partner, partner_roster, partner_taxi)
    
    success, msg = github_api_request("data/rosters.json", "PUT", {
        "message": f"Trade executed: {proposer} <-> {partner}",
        "content": rosters
    })
    
    if not success:
        return False, f"Failed to save rosters: {msg}"
    
    return True, "Trade executed successfully"


def handle_respond_trade(data: dict) -> tuple[int, dict]:
    """Handle trade acceptance or rejection."""
    team = data.get("team")
    password = data.get("password")
    trade_id = data.get("trade_id")
    accept = data.get("accept", False)
    
    valid, msg = validate_team(team, password)
    if not valid:
        return 401, {"error": msg}
    
    if not trade_id:
        return 400, {"error": "Missing trade_id"}
    
    success, result = github_api_request("data/pending_trades.json")
    if not success:
        return 500, {"error": result}
    
    pending = result["content"]
    trade = next((t for t in pending["trades"] if t["id"] == trade_id), None)
    
    if not trade:
        return 400, {"error": "Trade not found"}
    
    if trade["partner"] != team:
        return 403, {"error": "You are not the trade partner"}
    
    if trade["status"] != "pending":
        return 400, {"error": f"Trade is already {trade['status']}"}
    
    if accept:
        success, exec_msg = execute_trade(trade)
        if not success:
            return 500, {"error": exec_msg}
        
        trade["status"] = "accepted"
        trade["accepted_at"] = datetime.utcnow().isoformat()
        
        add_transaction_log({
            "type": "trade",
            "proposer": trade["proposer"],
            "partner": trade["partner"],
            "proposer_gives": trade["proposer_gives"],
            "proposer_receives": trade["proposer_receives"],
            "week": trade["week"],
            "timestamp": datetime.utcnow().isoformat()
        })
        
        message = "Trade accepted and executed"
    else:
        trade["status"] = "rejected"
        trade["rejected_at"] = datetime.utcnow().isoformat()
        message = "Trade rejected"
    
    success, msg = github_api_request("data/pending_trades.json", "PUT", {
        "message": f"Trade {trade_id} {trade['status']}",
        "content": pending
    })
    
    if not success:
        return 500, {"error": msg}
    
    return 200, {"success": True, "message": message}


def add_transaction_log(transaction: dict):
    """Add a transaction to the transaction log JSON file."""
    try:
        success, result = github_api_request("data/transaction_log.json")
        if success:
            log = result["content"]
        else:
            log = {"transactions": []}
        
        log["transactions"].append(transaction)
        
        success, msg = github_api_request("data/transaction_log.json", "PUT", {
            "message": f"Transaction logged: {transaction.get('type', 'unknown')}",
            "content": log
        })
        
        if not success:
            print(f"Failed to save transaction log: {msg}")
    except Exception as e:
        print(f"Error logging transaction: {e}")


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", "0")
        self.end_headers()
    
    def do_GET(self):
        """Handle GET requests."""
        self._send_json(200, {"status": "OPFL Transaction API is running"})
    
    def do_POST(self):
        """Handle transaction requests."""
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            data = json.loads(body.decode()) if body else {}
            
            action = data.get("action")
            
            if action == "validate":
                valid, msg = validate_team(data.get("team"), data.get("password"))
                if valid:
                    return self._send_json(200, {"success": True})
                else:
                    return self._send_json(401, {"error": msg})
            
            elif action == "taxi_activate":
                status, result = handle_taxi_activation(data)
                return self._send_json(status, result)
            
            elif action == "propose_trade":
                status, result = handle_propose_trade(data)
                return self._send_json(status, result)
            
            elif action == "respond_trade":
                status, result = handle_respond_trade(data)
                return self._send_json(status, result)
            
            else:
                return self._send_json(400, {"error": f"Unknown action: {action}"})
                
        except json.JSONDecodeError:
            return self._send_json(400, {"error": "Invalid JSON"})
        except Exception as e:
            return self._send_json(500, {"error": str(e)})
    
    def _send_json(self, status_code: int, data: dict):
        """Send JSON response with CORS headers."""
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())
    
    def log_message(self, format, *args):
        """Suppress default logging."""
        pass
