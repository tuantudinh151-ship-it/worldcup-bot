import os, json, logging, asyncio, io
from datetime import datetime, timedelta
import pytz, requests
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import Application, CommandHandler, PollAnswerHandler, ContextTypes
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ============================================================
#  ĐỌC CẤU HÌNH TỪ FILE .env (không cần sửa file này)
# ============================================================
def load_env():
    env = {}
    if os.path.exists(".env"):
        with open(".env", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    env[k.strip()] = v.strip()
    return env

_env = load_env()
BOT_TOKEN    = _env.get("BOT_TOKEN", "")
GROUP_ID     = int(_env.get("GROUP_ID", "0"))
ADMIN_ID     = int(_env.get("ADMIN_ID", "0"))
ODDS_API_KEY    = _env.get("ODDS_API_KEY", "")
APIFOOTBALL_KEY = _env.get("APIFOOTBALL_KEY", "")
TIMEZONE        = "Asia/Ho_Chi_Minh"

if not BOT_TOKEN or not GROUP_ID or not ADMIN_ID:
    print("LỖI: Chưa có file .env hoặc thiếu thông tin!")
    print("Hãy chạy setup_config.bat trước.")
    exit(1)
# ============================================================

FEE          = {"group": 50000, "knockout": 100000, "final": 200000}
SPORT_KEY    = "soccer_fifa_world_cup"
POLL_BEFORE  = 8   # Gửi poll trước kickoff bao nhiêu tiếng
TZ           = pytz.timezone(TIMEZONE)
DATA_FILE    = "data.json"
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============================================================
#  THÔNG TIN VÒNG ĐẤU (World Cup 2026)
# ============================================================
def get_round_type(commence_time: datetime) -> str:
    """
    World Cup 2026: 12/06 - 20/07
    Vòng bảng:         12/06 - 27/06 (48 trận)
    Vòng loại trực tiếp: 29/06 - 19/07
    Chung kết:         20/07
    """
    from datetime import date
    d = commence_time.astimezone(TZ).date()
    if d <= date(2026, 6, 27): return "group"     # Vòng bảng
    if d <= date(2026, 7, 19): return "knockout"  # Vòng loại trực tiếp
    return "final"                                 # Chung kết 20/07

# ============================================================
#  DATA
# ============================================================
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"matches": {}, "predictions": {}, "players": {}}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def is_admin(uid): return uid == ADMIN_ID

# ============================================================
#  LẤY KÈO TỪ THE ODDS API
# ============================================================
def fetch_odds(market="spreads"):
    try:
        r = requests.get(
            f"https://api.the-odds-api.com/v4/sports/{SPORT_KEY}/odds",
            params={"apiKey": ODDS_API_KEY, "regions": "eu",
                    "markets": market, "oddsFormat": "decimal", "dateFormat": "iso"},
            timeout=15
        )
        if r.status_code != 200:
            logger.error(f"Odds API lỗi: {r.status_code} {r.text}")
            return []
        return r.json()
    except Exception as e:
        logger.error(f"fetch_odds lỗi: {e}")
        return []

def parse_asian_handicap(game: dict) -> dict | None:
    try:
        priority = ["pinnacle", "betfair_ex_eu", "betonlineag", "mybookieag"]
        chosen_bm = None
        for bm_key in priority:
            for bm in game.get("bookmakers", []):
                if bm["key"] == bm_key:
                    chosen_bm = bm
                    break
            if chosen_bm: break
        if not chosen_bm and game.get("bookmakers"):
            chosen_bm = game["bookmakers"][0]
        if not chosen_bm: return None

        market = next((m for m in chosen_bm["markets"] if m["key"] == "spreads"), None)
        if not market: return None
        outcomes = market["outcomes"]
        if len(outcomes) < 2: return None

        home = next((o for o in outcomes if o["name"] == game["home_team"]), outcomes[0])
        away = next((o for o in outcomes if o["name"] == game["away_team"]), outcomes[1])

        return {
            "home_team":     game["home_team"],
            "away_team":     game["away_team"],
            "commence_time": game["commence_time"],
            "match_id":      game["id"][:8].upper(),
            "bookmaker":     chosen_bm["title"],
            "handicap":      home.get("point", 0),
            "home_odds":     home["price"],
            "away_odds":     away["price"],
        }
    except Exception as e:
        logger.error(f"parse lỗi: {e}")
        return None

def normalize_handicap(handicap: float) -> float:
    """
    Làm tròn kèo:
    0.25 → 0.5 | 0.75 → 0.5 | 1.25 → 1.5 | 1.75 → 1.5 | v.v.
    Giữ nguyên: 0, 0.5, 1, 1.5, 2...
    """
    sign  = -1 if handicap < 0 else 1
    abs_h = abs(handicap)
    frac  = round(abs_h % 1, 4)
    if frac in (0.25, 0.75):
        abs_h = int(abs_h) + 0.5
    return round(sign * abs_h, 2)

def has_draw_option(handicap: float) -> bool:
    """Kèo nguyên (0, 1, 2...) → có lựa chọn HÒA"""
    h = normalize_handicap(handicap)
    return h % 1 == 0

def format_handicap(handicap: float, home: str, away: str) -> str:
    h = normalize_handicap(handicap)
    if h == 0:        return "Kèo chẵn 0 (có thêm HÒA)"
    elif h < 0:       return f"{home} chấp {abs(h)} trái"
    else:             return f"{away} chấp {h} trái"

def build_poll_options(m: dict) -> list[str]:
    """
    Kèo nguyên (0, 1, 2...): 3 lựa chọn TRÊN / HÒA / DƯỚI
    Kèo lẻ (0.5, 1.5...): 2 lựa chọn TRÊN / DƯỚI
    Kèo 0.25/0.75 đã làm tròn → 0.5 → 2 lựa chọn
    Sai = mất 100% trong mọi trường hợp.
    """
    h    = normalize_handicap(m["handicap"])
    home = m["home_team"]
    away = m["away_team"]
    ho   = m["home_odds"]
    ao   = m["away_odds"]

    if h % 1 == 0:
        if h == 0:
            return [
                f"TRÊN - {home} thắng ({ho})",
                f"HÒA - Trận đấu hòa",
                f"DƯỚI - {away} thắng ({ao})",
            ]
        else:
            return [
                f"TRÊN - {home} ({ho})",
                f"HÒA KÈO - Trận hòa đúng kèo",
                f"DƯỚI - {away} ({ao})",
            ]
    else:
        return [
            f"TRÊN - {home} ({ho})",
            f"DƯỚI - {away} ({ao})",
        ]

def get_choice_from_index(index: int, m: dict) -> str:
    """Chuyển index bình chọn → home/draw/away"""
    if has_draw_option(m.get("handicap", 0)):
        return ["home", "draw", "away"][index] if index < 3 else "away"
    else:
        return "home" if index == 0 else "away"

# ============================================================
#  GỬI VÀ KHÓA POLL
# ============================================================

async def auto_send_poll_job(match_id: str, app):
    """Gửi poll vào nhóm và ghim lên đầu"""
    data = load_data()
    if match_id not in data["matches"]:
        return
    m = data["matches"][match_id]
    if m.get("poll_message_id"):
        return  # Đã gửi rồi

    kickoff = datetime.fromisoformat(m["kickoff"])
    fee = FEE.get(m["round"], 50000)
    hcap_text = format_handicap(m["handicap"], m["home_team"], m["away_team"])
    options = build_poll_options(m)
    even_note = "\nLưu ý: Kèo chẵn - có lựa chọn HÒA!" if has_draw_option(m["handicap"]) else ""

    question = (
        f"BÌNH CHỌN {match_id}: {m['home_team']} vs {m['away_team']}\n"
        f"Kickoff: {kickoff.strftime('%d/%m/%Y %H:%M')} | Phí thua: {fee:,}đ\n"
        f"Kèo: {hcap_text}{even_note}"
    )

    try:
        # Poll ẨN DANH: thành viên không thấy ai vote gì, không thấy số phiếu
        # cho đến khi poll đóng (giới hạn của Telegram với anonymous poll)
        msg = await app.bot.send_poll(
            chat_id=GROUP_ID,
            question=question,
            options=options,
            is_anonymous=False,            # Công khai - thấy ai bình chọn gì
            allows_multiple_answers=False,
            open_period=None               # Không tự đóng, bot khóa đúng kickoff
        )

        # Ghim poll lên đầu nhóm
        await app.bot.pin_chat_message(
            chat_id=GROUP_ID,
            message_id=msg.message_id,
            disable_notification=True   # Ghim không thông báo ồn ào
        )

        data["matches"][match_id]["poll_id"] = msg.poll.id
        data["matches"][match_id]["poll_message_id"] = msg.message_id
        data["matches"][match_id]["has_draw_option"] = has_draw_option(m["handicap"])
        save_data(data)
        logger.info(f"Đã gửi + ghim poll trận {match_id}")
    except Exception as e:
        logger.error(f"Lỗi gửi poll {match_id}: {e}")


async def lock_poll_job(match_id: str, app):
    """Khóa poll + bỏ ghim khi kickoff"""
    data = load_data()
    if match_id not in data["matches"]:
        return
    m = data["matches"][match_id]
    if m["locked"] or not m.get("poll_message_id"):
        return
    try:
        # Khóa poll
        await app.bot.stop_poll(chat_id=GROUP_ID, message_id=m["poll_message_id"])

        # Bỏ ghim
        await app.bot.unpin_chat_message(
            chat_id=GROUP_ID,
            message_id=m["poll_message_id"]
        )

        data["matches"][match_id]["locked"] = True
        save_data(data)

        await app.bot.send_message(
            chat_id=GROUP_ID,
            text=(
                f"KHÓA BÌNH CHỌN - Trận {match_id}\n"
                f"{m['home_team']} vs {m['away_team']}\n"
                f"Bóng lăn rồi! Chờ kết quả..."
            )
        )
        logger.info(f"Đã khóa + bỏ ghim poll trận {match_id}")
    except Exception as e:
        logger.error(f"Lỗi khóa poll {match_id}: {e}")


# ============================================================
#  LỆNH ADMIN
# ============================================================

async def cmd_newmatch(update, context):
    if not is_admin(update.effective_user.id): return
    args = context.args
    if len(args) < 9:
        await update.message.reply_text(
            'Cu phap:\n'
            '/newmatch <id> <doi1> vs <doi2> <ngay> <gio> <vong> <handicap> <odds_tren> <odds_duoi>\n'
            'Vi du:\n'
            '/newmatch TEST Brazil vs Argentina 2026-06-07 22:05 group -0.5 1.95 1.95\n'
            'handicap: -0.5=nha chap nua | 0=keo chan | 1=khach chap 1'
        )
        return
    try:
        match_id   = args[0].upper()
        vs_idx     = args.index('vs')
        team1      = ' '.join(args[1:vs_idx])
        team2      = ' '.join(args[vs_idx+1:-6])
        date_str   = args[-6]
        time_str   = args[-5]
        round_type = args[-4]
        handicap   = float(args[-3])
        home_odds  = float(args[-2])
        away_odds  = float(args[-1])
        if round_type not in ['group', 'knockout', 'final']:
            raise ValueError('Vong phai la: group / knockout / final')
        kickoff = TZ.localize(datetime.strptime(date_str + ' ' + time_str, '%Y-%m-%d %H:%M'))
        data = load_data()
        if match_id in data['matches']:
            await update.message.reply_text('Tran ' + match_id + ' da ton tai.')
            return
        h_norm = normalize_handicap(handicap)
        data['matches'][match_id] = {
            'home_team': team1, 'away_team': team2,
            'kickoff': kickoff.isoformat(), 'round': round_type,
            'handicap': h_norm, 'home_odds': home_odds, 'away_odds': away_odds,
            'bookmaker': 'Thu cong', 'has_draw_option': has_draw_option(h_norm),
            'result': None, 'poll_id': None, 'poll_message_id': None, 'locked': False
        }
        save_data(data)
        scheduler = context.bot_data.get('scheduler')
        now = datetime.now(TZ)
        send_time = kickoff - timedelta(hours=POLL_BEFORE)
        if send_time > now and scheduler:
            scheduler.add_job(auto_send_poll_job, 'date', run_date=send_time,
                args=[match_id, context.application], id='send_' + match_id, replace_existing=True)
        if kickoff > now and scheduler:
            scheduler.add_job(lock_poll_job, 'date', run_date=kickoff,
                args=[match_id, context.application], id='lock_' + match_id, replace_existing=True)
        fee = FEE.get(round_type, 50000)
        hcap_txt = format_handicap(h_norm, team1, team2)
        await update.message.reply_text(
            'Da tao tran ' + match_id + ':\n' +
            team1 + ' vs ' + team2 + '\n' +
            'Kickoff: ' + kickoff.strftime('%d/%m/%Y %H:%M') + '\n' +
            'Keo: ' + hcap_txt + ' | Phi thua: ' + str(fee) + 'd\n' +
            'Go /sendpoll ' + match_id + ' de gui poll vao nhom.'
        )
    except Exception as e:
        await update.message.reply_text('Loi: ' + str(e))


async def cmd_fetchodds(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xem kèo mới từ API. /fetchodds"""
    if not is_admin(update.effective_user.id): return
    await update.message.reply_text("Đang lấy kèo...")

    games = fetch_odds("spreads")
    if not games:
        await update.message.reply_text("Không lấy được dữ liệu. Kiểm tra API key.")
        return

    now = datetime.now(TZ)
    upcoming = []
    for g in games:
        p = parse_asian_handicap(g)
        if not p: continue
        kickoff = datetime.fromisoformat(p["commence_time"].replace("Z", "+00:00")).astimezone(TZ)
        if kickoff > now:
            upcoming.append((kickoff, p))
    upcoming.sort(key=lambda x: x[0])

    if not upcoming:
        await update.message.reply_text("Không có trận nào sắp tới.")
        return

    msg = f"DANH SÁCH TRẬN ({len(upcoming)} trận)\n/autosetup để tự động tạo tất cả\n" + "="*32 + "\n"
    for i, (kickoff, p) in enumerate(upcoming[:12], 1):
        hcap = format_handicap(p["handicap"], p["home_team"], p["away_team"])
        even = " [CHẴN+HÒA]" if has_draw_option(p["handicap"]) else ""
        msg += (
            f"{i}. {p['home_team']} vs {p['away_team']}\n"
            f"   {kickoff.strftime('%d/%m %H:%M')} | {hcap}{even}\n"
            f"   Tỷ lệ: {p['home_odds']} / {p['away_odds']} | {p['bookmaker']}\n\n"
        )
    await update.message.reply_text(msg)


async def cmd_autosetup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Tự động tạo trận + lên lịch gửi poll và khóa. /autosetup [số_trận]"""
    if not is_admin(update.effective_user.id): return
    limit = int(context.args[0]) if context.args else 999
    await update.message.reply_text(f"Đang lấy kèo và tạo lịch...")

    games = fetch_odds("spreads")
    if not games:
        await update.message.reply_text("Không lấy được dữ liệu API.")
        return

    data = load_data()
    now = datetime.now(TZ)
    scheduler = context.bot_data.get("scheduler")
    created = skipped = 0

    for g in sorted(games, key=lambda x: x["commence_time"]):
        if created >= limit: break
        p = parse_asian_handicap(g)
        if not p: continue

        kickoff = datetime.fromisoformat(p["commence_time"].replace("Z", "+00:00")).astimezone(TZ)
        if kickoff <= now: continue

        match_id = p["match_id"]
        base_id = match_id
        counter = 1
        while match_id in data["matches"] and data["matches"][match_id]["home_team"] != p["home_team"]:
            match_id = f"{base_id}{counter}"
            counter += 1

        if match_id in data["matches"]:
            skipped += 1
            continue

        round_type = get_round_type(kickoff)
        data["matches"][match_id] = {
            "home_team":      p["home_team"],
            "away_team":      p["away_team"],
            "kickoff":        kickoff.isoformat(),
            "round":          round_type,
            "handicap":       p["handicap"],
            "home_odds":      p["home_odds"],
            "away_odds":      p["away_odds"],
            "bookmaker":      p["bookmaker"],
            "has_draw_option": has_draw_option(p["handicap"]),
            "result":         None,
            "poll_id":        None,
            "poll_message_id": None,
            "locked":         False
        }

        # Lên lịch gửi poll trước POLL_BEFORE tiếng
        send_time = kickoff - timedelta(hours=POLL_BEFORE)
        if send_time > now and scheduler:
            scheduler.add_job(auto_send_poll_job, "date", run_date=send_time,
                args=[match_id, context.application],
                id=f"send_{match_id}", replace_existing=True)

        # Lên lịch khóa + bỏ ghim lúc kickoff
        if scheduler:
            scheduler.add_job(lock_poll_job, "date", run_date=kickoff,
                args=[match_id, context.application],
                id=f"lock_{match_id}", replace_existing=True)

        created += 1

    save_data(data)
    await update.message.reply_text(
        f"Đã tạo {created} trận, bỏ qua {skipped} trận đã có.\n"
        f"Poll tự gửi trước {POLL_BEFORE} tiếng, ghim lên đầu nhóm.\n"
        f"Poll tự khóa + bỏ ghim đúng giờ kickoff.\n"
        f"Dùng /matches để xem danh sách."
    )


async def cmd_sendpoll(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Gửi poll thủ công. /sendpoll <mã_trận>"""
    if not is_admin(update.effective_user.id): return
    if not context.args:
        await update.message.reply_text("Cú pháp: /sendpoll <mã_trận>")
        return
    match_id = context.args[0].upper()
    await auto_send_poll_job(match_id, context.application)
    await update.message.reply_text(f"Đã gửi + ghim poll trận {match_id}.")


async def cmd_lockpoll(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Khóa poll thủ công. /lockpoll <mã_trận>"""
    if not is_admin(update.effective_user.id): return
    if not context.args:
        await update.message.reply_text("Cú pháp: /lockpoll <mã_trận>")
        return
    match_id = context.args[0].upper()
    await lock_poll_job(match_id, context.application)
    await update.message.reply_text(f"Đã khóa + bỏ ghim poll trận {match_id}.")


async def cmd_result(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Nhập kết quả kèo. /result <mã_trận> <home|draw|away>
    home  = đội nhà thắng kèo (TRÊN)
    draw  = hòa kèo (chỉ với kèo chẵn, hoàn 50%)
    away  = đội khách thắng kèo (DƯỚI)
    """
    if not is_admin(update.effective_user.id): return
    if len(context.args) < 2:
        await update.message.reply_text("Cú pháp: /result <mã_trận> <home|draw|away>")
        return

    match_id = context.args[0].upper()
    result   = context.args[1].lower()
    if result not in ["home", "draw", "away"]:
        await update.message.reply_text("Kết quả phải là: home / draw / away")
        return

    data = load_data()
    if match_id not in data["matches"]:
        await update.message.reply_text(f"Không tìm thấy trận {match_id}.")
        return

    m = data["matches"][match_id]
    data["matches"][match_id]["result"] = result
    fee  = FEE.get(m["round"], 50000)

    winners, losers = [], []
    for user_id, pred in data["predictions"].get(match_id, {}).items():
        player = data["players"].setdefault(user_id, {"name": pred["name"], "debt": 0})
        player["name"] = pred["name"]
        choice = pred["choice"]

        if choice == result:
            winners.append(pred["name"])
        else:
            # Sai = mất 100% trong mọi trường hợp
            losers.append(pred["name"])
            player["debt"] = player.get("debt", 0) + fee

    save_data(data)

    result_label = {"home": m["home_team"], "draw": "Hòa kèo", "away": m["away_team"]}[result]
    hcap_text = format_handicap(m["handicap"], m["home_team"], m["away_team"])

    msg = (
        f"KẾT QUẢ TRẬN {match_id}\n"
        f"{m['home_team']} vs {m['away_team']}\n"
        f"Kèo: {hcap_text}\n"
        f"Kết quả kèo: {result_label}\n\n"
    )
    if winners: msg += f"THẮNG ({len(winners)}): {', '.join(winners)}\n"
    if losers:  msg += f"THUA -{fee:,}đ ({len(losers)}): {', '.join(losers)}\n"
    if not data["predictions"].get(match_id):
        msg += "Không có ai bình chọn trận này."

    await context.bot.send_message(chat_id=GROUP_ID, text=msg)
    await update.message.reply_text("Đã cập nhật kết quả.")


# ============================================================
#  LỆNH CHO TẤT CẢ
# ============================================================

async def cmd_standings(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["players"]:
        await update.message.reply_text("Chưa có dữ liệu.")
        return
    sorted_p = sorted(data["players"].items(), key=lambda x: x[1].get("debt", 0), reverse=True)
    msg = "BẢNG XẾP HẠNG NỢ\n" + "="*28 + "\n"
    for i, (uid, p) in enumerate(sorted_p, 1):
        debt = p.get("debt", 0)
        icon = "🔴" if debt > 0 else "🟢"
        msg += f"{i}. {icon} {p['name']}: {'-' if debt>0 else ''}{debt:,}đ\n"
    await update.message.reply_text(msg)


async def cmd_matches(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    now = datetime.now(TZ)
    upcoming = [
        (mid, m, datetime.fromisoformat(m["kickoff"]))
        for mid, m in data["matches"].items()
        if datetime.fromisoformat(m["kickoff"]) > now and not m.get("result")
    ]
    upcoming.sort(key=lambda x: x[2])
    if not upcoming:
        await update.message.reply_text("Không có trận nào sắp tới.")
        return
    msg = f"CÁC TRẬN SẮP TỚI ({len(upcoming)} trận)\n" + "="*28 + "\n"
    for mid, m, kickoff in upcoming[:10]:
        status = "📌 Đã ghim" if m.get("poll_message_id") and not m["locked"] else ("🔒 Đã khóa" if m["locked"] else "⏳ Chờ gửi")
        hcap = format_handicap(m.get("handicap", 0), m["home_team"], m["away_team"])
        even = " [+HÒA]" if m.get("has_draw_option") else ""
        msg += f"{mid}: {m['home_team']} vs {m['away_team']}\n"
        msg += f"   {kickoff.strftime('%d/%m %H:%M')} | {status} | {hcap}{even}\n"
    await update.message.reply_text(msg)


async def cmd_mypredictions(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = str(update.effective_user.id)
    data = load_data()
    history = []
    for mid, preds in data["predictions"].items():
        if user_id in preds:
            m = data["matches"].get(mid, {})
            result = m.get("result")
            pred   = preds[user_id]["choice"]
            label  = {"home": m.get("home_team","?"), "draw": "Hòa", "away": m.get("away_team","?")}.get(pred, pred)
            outcome = ("THẮNG" if pred == result else "THUA") if result else "Chờ kết quả"
            history.append(f"{mid}: Chọn {label} → {outcome}")
    if not history:
        await update.message.reply_text("Bạn chưa bình chọn trận nào.")
        return
    debt = data["players"].get(user_id, {}).get("debt", 0)
    msg = "LỊCH SỬ BÌNH CHỌN\n" + "\n".join(history) + f"\n\nTổng nợ: {debt:,}đ"
    await update.message.reply_text(msg)


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "HƯỚNG DẪN BOT WORLD CUP 2026\n" + "="*30 + "\n\n"
        "LỆNH TẤT CẢ:\n"
        "/join — Đăng ký tham gia\n""/matches — Trận sắp tới\n"
        "/standings — Bảng nợ tiền\n"
        "/mypredictions — Lịch sử của bạn\n\n"
        "LỆNH ADMIN:\n"
        "/newmatch — Tạo trận thủ công (test)\n""/fetchodds — Xem kèo mới từ nhà cái\n"
        "/autosetup [số] — Tạo trận + lên lịch tự động\n"
        "/sendpoll <id> — Gửi + ghim poll thủ công\n"
        "/lockpoll <id> — Khóa + bỏ ghim thủ công\n"
        "/result <id> <home|draw|away> — Nhập kết quả\n\n"
        "PHÍ THUA:\n"
        "Vòng bảng: 50,000đ | KO: 100,000đ | CK: 200,000đ\n\n"
        "KÈO CHÂU Á:\n"
        "TRÊN = đội chấp thắng kèo\n"
        "DƯỚI = đội được chấp thắng kèo\n"
        "HÒA / HÒA KÈO = chỉ xuất hiện ở kèo nguyên (0, 1, 2...)\n"
        "  Sai = mất 100% trong mọi trường hợp\n\n"
        "BÌNH CHỌN CÔNG KHAI:\n"
        "  - Mọi người thấy ai bình chọn gì\n"
        "  - Có thể đổi ý cho đến khi poll khóa\n"
        "  - Có thể đổi ý cho đến khi poll khóa!"
    )


# ============================================================
#  XỬ LÝ BÌNH CHỌN (ẨN DANH)
# ============================================================

async def handle_poll_answer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Lưu bình chọn ẨN DANH - chỉ bot nhận được thông tin, nhóm không biết ai chọn gì.
    Người dùng CHỈ ĐƯỢC CHỌN 1 LẦN - không đổi ý được.
    Số phiếu bầu ẩn cho đến khi poll đóng (tính năng mặc định của Telegram anonymous poll).
    """
    answer  = update.poll_answer
    poll_id = answer.poll_id
    user    = answer.user
    user_id = str(user.id)

    # Người dùng bỏ vote -> xóa bình chọn cũ
    if not answer.option_ids:
        data = load_data()
        match_id = next((mid for mid, m in data["matches"].items() if m.get("poll_id") == poll_id), None)
        if match_id:
            data["predictions"].get(match_id, {}).pop(user_id, None)
            save_data(data)
        return

    option_index = answer.option_ids[0]
    data = load_data()
    match_id = next((mid for mid, m in data["matches"].items() if m.get("poll_id") == poll_id), None)
    if not match_id or data["matches"][match_id]["locked"]:
        return

    m = data["matches"][match_id]

    choice = get_choice_from_index(option_index, m)
    # Cho phép đổi ý - ghi đè bình chọn cũ
    data["predictions"].setdefault(match_id, {})[user_id] = {
        "name": user.full_name, "choice": choice
    }
    data["players"].setdefault(user_id, {"name": user.full_name, "debt": 0})["name"] = user.full_name
    save_data(data)
    logger.info(f"{user.full_name} bình chọn '{choice}' cho trận {match_id}")


# ============================================================
#  KHỞI ĐỘNG
# ============================================================

async def post_init(application):
    scheduler = AsyncIOScheduler(timezone=TZ)
    scheduler.start()
    application.bot_data["scheduler"] = scheduler

    data = load_data()
    now  = datetime.now(TZ)
    restored = 0

    for match_id, m in data["matches"].items():
        kickoff = datetime.fromisoformat(m["kickoff"])

        # Khôi phục lịch gửi poll
        if not m.get("poll_message_id") and kickoff > now:
            send_time = kickoff - timedelta(hours=POLL_BEFORE)
            if send_time > now:
                scheduler.add_job(auto_send_poll_job, "date", run_date=send_time,
                    args=[match_id, application], id=f"send_{match_id}", replace_existing=True)
                restored += 1

        # Khôi phục lịch khóa poll
        if not m["locked"] and m.get("poll_message_id") and kickoff > now:
            scheduler.add_job(lock_poll_job, "date", run_date=kickoff,
                args=[match_id, application], id=f"lock_{match_id}", replace_existing=True)
            restored += 1

    logger.info(f"Bot đã khởi động! Khôi phục {restored} lịch.")


# ============================================================
#  LẤY KẾT QUẢ TỪ API-FOOTBALL
# ============================================================

def fetch_match_result_api(home_team: str, away_team: str, match_date: str) -> dict | None:
    """
    Lấy kết quả trận đấu từ API-Football.
    match_date: định dạng YYYY-MM-DD
    Trả về: {"home_score": int, "away_score": int} hoặc None nếu chưa có
    """
    if not APIFOOTBALL_KEY:
        return None
    try:
        # FIFA World Cup 2026 = league_id 1 (FIFA WC), season 2026
        url = "https://v3.football.api-sports.io/fixtures"
        headers = {"x-apisports-key": APIFOOTBALL_KEY}
        params  = {
            "league": "1",       # FIFA World Cup
            "season": "2026",
            "date":   match_date,
            "status": "FT"       # Chỉ lấy trận đã kết thúc
        }
        r = requests.get(url, headers=headers, params=params, timeout=15)
        if r.status_code != 200:
            logger.error(f"API-Football lỗi: {r.status_code}")
            return None

        fixtures = r.json().get("response", [])
        for fix in fixtures:
            h = fix["teams"]["home"]["name"]
            a = fix["teams"]["away"]["name"]
            # So khớp tên đội (linh hoạt)
            if (home_team.lower() in h.lower() or h.lower() in home_team.lower()) and \
               (away_team.lower() in a.lower() or a.lower() in away_team.lower()):
                goals = fix["goals"]
                return {
                    "home_score": goals["home"],
                    "away_score": goals["away"]
                }
        return None
    except Exception as e:
        logger.error(f"fetch_match_result_api lỗi: {e}")
        return None


def determine_keo_result(home_score: int, away_score: int, handicap: float) -> str:
    """
    Tính kết quả kèo châu Á từ tỷ số thực tế.
    handicap < 0: đội nhà chấp |handicap| trái
    handicap > 0: đội khách chấp handicap trái
    handicap == 0: kèo chẵn, có thể hòa
    Trả về: 'home' | 'draw' | 'away'
    """
    h = normalize_handicap(handicap)
    # Tính tỷ số sau khi áp kèo
    # handicap âm = đội nhà chấp: home_score + handicap (vd: -0.5 → home cần thắng)
    adjusted_home = home_score + h  # h âm = nhà chấp, dương = khách chấp

    if adjusted_home > away_score:
        return "home"
    elif adjusted_home < away_score:
        return "away"
    else:
        return "draw"  # chỉ xảy ra với kèo nguyên


async def cmd_fetchresult(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Tự động lấy kết quả từ API-Football cho tất cả trận đã kết thúc.
    /fetchresult
    """
    if not is_admin(update.effective_user.id): return
    if not APIFOOTBALL_KEY:
        await update.message.reply_text("Chưa có APIFOOTBALL_KEY trong file .env!")
        return

    await update.message.reply_text("Đang lấy kết quả từ API-Football...")
    data = load_data()
    now  = datetime.now(TZ)
    updated = 0
    errors  = 0

    for match_id, m in data["matches"].items():
        if m.get("result"):
            continue  # Đã có kết quả rồi
        kickoff = datetime.fromisoformat(m["kickoff"])
        # Chỉ xử lý trận đã kết thúc hơn 2 tiếng
        if now < kickoff + timedelta(hours=2):
            continue

        match_date = kickoff.strftime("%Y-%m-%d")
        score = fetch_match_result_api(m["home_team"], m["away_team"], match_date)
        if not score:
            errors += 1
            continue

        keo_result = determine_keo_result(score["home_score"], score["away_score"], m["handicap"])
        data["matches"][match_id]["result"]     = keo_result
        data["matches"][match_id]["home_score"] = score["home_score"]
        data["matches"][match_id]["away_score"] = score["away_score"]

        # Tính thắng/thua cho tất cả thành viên
        fee       = FEE.get(m["round"], 50000)
        members   = data.get("members", {})
        preds     = data["predictions"].get(match_id, {})
        winners, losers, no_vote = [], [], []

        for uid, member in members.items():
            if uid in preds:
                choice = preds[uid]["choice"]
                player = data["players"].setdefault(uid, {"name": member["name"], "debt": 0})
                player["name"] = member["name"]
                if choice == keo_result:
                    winners.append(member["name"])
                else:
                    losers.append(member["name"])
                    player["debt"] = player.get("debt", 0) + fee
            else:
                # Không bình chọn = thua cuộc
                player = data["players"].setdefault(uid, {"name": member["name"], "debt": 0})
                player["name"] = member["name"]
                no_vote.append(member["name"])
                player["debt"] = player.get("debt", 0) + fee

        save_data(data)
        updated += 1

        # Thông báo kết quả vào nhóm
        hcap_txt     = format_handicap(m["handicap"], m["home_team"], m["away_team"])
        result_label = {"home": m["home_team"], "draw": "Hòa kèo", "away": m["away_team"]}[keo_result]
        msg = (
            f"KẾT QUẢ TRẬN {match_id}\n"
            f"{m['home_team']} {score['home_score']} - {score['away_score']} {m['away_team']}\n"
            f"Kèo: {hcap_txt} → {result_label} thắng kèo\n\n"
        )
        if winners:  msg += f"THẮNG ({len(winners)}): {', '.join(winners)}\n"
        if losers:   msg += f"THUA -{fee:,}đ ({len(losers)}): {', '.join(losers)}\n"
        if no_vote:  msg += f"KHÔNG BQ -{fee:,}đ ({len(no_vote)}): {', '.join(no_vote)}\n"

        try:
            await context.bot.send_message(chat_id=GROUP_ID, text=msg)
        except Exception as e:
            logger.error(f"Lỗi gửi kết quả {match_id}: {e}")

    await update.message.reply_text(
        f"Đã cập nhật {updated} trận.\n"
        f"Không tìm thấy kết quả: {errors} trận.\n"
        f"Gõ /excel để xuất file tính tiền."
    )


# ============================================================
#  QUẢN LÝ THÀNH VIÊN
# ============================================================


async def cmd_join(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Thành viên tự đăng ký tham gia. Gõ /join trong nhóm."""
    user    = update.effective_user
    user_id = str(user.id)
    name    = user.full_name
    data    = load_data()

    if user_id in data.get("members", {}):
        await update.message.reply_text(
            f"Bạn đã đăng ký rồi: {name}\nGõ /mypredictions để xem lịch sử bình chọn."
        )
        return

    data.setdefault("members", {})[user_id] = {"name": name, "paid": 0}
    data["players"].setdefault(user_id, {"name": name, "debt": 0})["name"] = name
    save_data(data)

    # Đếm số thành viên
    total = len(data["members"])
    await update.message.reply_text(
        f"Đã đăng ký tham gia!\n"
        f"Tên: {name}\n"
        f"Tổng thành viên: {total} người\n\n"
        f"Bạn sẽ bị tính thua nếu không bình chọn trước giờ kickoff."
    )
    # Thông báo cho admin
    try:
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"Thành viên mới đăng ký:\n{name} (ID: {user_id})\nTổng: {total} người"
        )
    except Exception:
        pass


async def cmd_members(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xem danh sách thành viên đã đăng ký. /members"""
    if not is_admin(update.effective_user.id): return
    data    = load_data()
    members = data.get("members", {})
    if not members:
        await update.message.reply_text(
            "Chưa có thành viên nào.\n"
            "Nhờ mọi người gõ /join trong nhóm để đăng ký."
        )
        return
    msg = f"DANH SÁCH THÀNH VIÊN ({len(members)} người)\n" + "="*28 + "\n"
    for i, (uid, m) in enumerate(members.items(), 1):
        debt   = data["players"].get(uid, {}).get("debt", 0)
        paid   = m.get("paid", 0)
        remain = max(0, debt - paid)
        status = f"Còn nợ {remain:,}đ" if remain > 0 else "Đã đóng đủ"
        msg   += f"{i}. {m['name']} — {status}\n"
    await update.message.reply_text(msg)


async def cmd_removemember(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xóa thành viên khỏi danh sách. /removemember <user_id>"""
    if not is_admin(update.effective_user.id): return
    if not context.args:
        await update.message.reply_text("Cú pháp: /removemember <user_id>")
        return
    uid  = str(context.args[0])
    data = load_data()
    if uid not in data.get("members", {}):
        await update.message.reply_text("Không tìm thấy thành viên này.")
        return
    name = data["members"][uid]["name"]
    del data["members"][uid]
    save_data(data)
    await update.message.reply_text(f"Đã xóa {name} khỏi danh sách.")


async def cmd_addmember(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Thêm thành viên vào danh sách tham gia.
    /addmember @username TênHiển_Thị
    hoặc /addmember 123456789 Tên
    """
    if not is_admin(update.effective_user.id): return
    if len(context.args) < 2:
        await update.message.reply_text(
            "Cú pháp: /addmember <user_id> <Tên>\n"
            "Ví dụ: /addmember 123456789 Nguyễn Văn A\n\n"
            "Lấy user_id: nhờ thành viên nhắn tin cho @userinfobot"
        )
        return
    try:
        uid  = str(context.args[0])
        name = " ".join(context.args[1:])
        data = load_data()
        data.setdefault("members", {})[uid] = {"name": name, "paid": 0}
        data["players"].setdefault(uid, {"name": name, "debt": 0})["name"] = name
        save_data(data)
        await update.message.reply_text(f"Đã thêm thành viên: {name} (ID: {uid})")
    except Exception as e:
        await update.message.reply_text(f"Lỗi: {e}")


async def cmd_paid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Xác nhận thành viên đã đóng tiền.
    /paid <tên hoặc user_id> <số_tiền>
    Ví dụ: /paid Nguyễn Văn A 150000
            /paid 123456789 150000
    """
    if not is_admin(update.effective_user.id): return
    if len(context.args) < 2:
        await update.message.reply_text(
            "Cú pháp: /paid <tên hoặc ID> <số_tiền>\n"
            "Ví dụ: /paid Nguyễn Văn A 150000\n"
            "        /paid 123456789 150000"
        )
        return
    try:
        # Số cuối luôn là số tiền
        amount_str = context.args[-1]
        # Loại bỏ dấu chấm/phẩy nếu có (ví dụ: 150.000 hoặc 150,000)
        amount_str = amount_str.replace(".", "").replace(",", "")
        if not amount_str.isdigit():
            await update.message.reply_text(
                "Số tiền phải là số nguyên ở cuối lệnh.\n"
                "Ví dụ: /paid Văn A 150000\n"
                "        /paid Trần Thị B 100000"
            )
            return
        amount = int(amount_str)
        query  = " ".join(context.args[:-1]).strip()
        data      = load_data()
        members   = data.get("members", {})

        # Tìm thành viên theo ID hoặc tên
        found_uid  = None
        found_name = None

        # Thử tìm theo ID trước
        if query in members:
            found_uid  = query
            found_name = members[query]["name"]
        else:
            # Tìm theo tên (không phân biệt hoa thường)
            query_lower = query.lower()
            matches_found = []
            for uid, m in members.items():
                if query_lower in m["name"].lower():
                    matches_found.append((uid, m["name"]))

            if len(matches_found) == 1:
                found_uid, found_name = matches_found[0]
            elif len(matches_found) > 1:
                names = "\n".join([f"• {n} (ID: {u})" for u, n in matches_found])
                await update.message.reply_text(
                    f"Tìm thấy {len(matches_found)} người tên gần giống:\n{names}\n\n"
                    "Dùng ID để xác nhận chính xác hơn:\n"
                    f"/paid <ID> {amount}"
                )
                return
            else:
                # Gợi ý danh sách thành viên
                all_names = "\n".join([f"• {m['name']} (ID: {u})" for u, m in members.items()])
                await update.message.reply_text(
                    f"Không tìm thấy '{query}'.\n\n"
                    f"Danh sách thành viên:\n{all_names}"
                )
                return

        # Cập nhật tiền đã đóng
        data["members"][found_uid]["paid"] = data["members"][found_uid].get("paid", 0) + amount
        player = data["players"].setdefault(found_uid, {"name": found_name, "debt": 0})
        debt   = player.get("debt", 0)
        paid   = data["members"][found_uid]["paid"]
        remain = max(0, debt - paid)
        save_data(data)

        status = "Đã đóng đủ!" if remain == 0 else f"Còn thiếu: {remain:,}đ"
        await update.message.reply_text(
            f"Đã ghi nhận {found_name} đóng {amount:,}đ\n"
            f"Tổng nợ: {debt:,}đ | Đã đóng: {paid:,}đ\n"
            f"{status}"
        )
    except Exception as e:
        await update.message.reply_text(f"Lỗi: {e}")


async def cmd_unpaid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xem danh sách chưa đóng đủ tiền. /unpaid"""
    if not is_admin(update.effective_user.id): return
    data    = load_data()
    members = data.get("members", {})
    if not members:
        await update.message.reply_text("Chưa có thành viên nào. Dùng /addmember để thêm.")
        return

    msg = "DANH SÁCH CHƯA ĐÓNG ĐỦ TIỀN\n" + "="*30 + "\n"
    total_owed = 0
    for uid, member in members.items():
        player = data["players"].get(uid, {})
        debt   = player.get("debt", 0)
        paid   = member.get("paid", 0)
        remain = debt - paid
        if remain > 0:
            msg += f"• {member['name']}: nợ {debt:,}đ | đã đóng {paid:,}đ | còn {remain:,}đ\n"
            total_owed += remain
    msg += f"\nTổng chưa thu: {total_owed:,}đ"
    await update.message.reply_text(msg)


# ============================================================
#  XUẤT EXCEL
# ============================================================

async def cmd_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xuất file Excel tính tiền cho tất cả thành viên. /excel"""
    if not is_admin(update.effective_user.id): return

    data    = load_data()
    members = data.get("members", {})
    matches = data.get("matches", {})

    if not members:
        await update.message.reply_text("Chưa có thành viên nào. Dùng /addmember để thêm.")
        return

    await update.message.reply_text("Đang tạo file Excel...")

    # Lấy danh sách trận đã có kết quả, sắp theo thời gian
    done_matches = {
        mid: m for mid, m in matches.items() if m.get("result")
    }
    done_matches = dict(sorted(done_matches.items(),
                               key=lambda x: x[1]["kickoff"]))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tính Tiền World Cup 2026"

    # ---- Style ----
    GREEN  = PatternFill("solid", fgColor="C6EFCE")
    RED    = PatternFill("solid", fgColor="FFC7CE")
    YELLOW = PatternFill("solid", fgColor="FFEB9C")
    GRAY   = PatternFill("solid", fgColor="D9D9D9")
    BLUE   = PatternFill("solid", fgColor="BDD7EE")
    HEADER = PatternFill("solid", fgColor="1F4E79")
    bold_white = Font(bold=True, color="FFFFFF")
    bold_black = Font(bold=True)
    center     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin       = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    # ---- Tiêu đề chính ----
    ws.merge_cells("A1:B1")
    ws["A1"] = "BẢNG TÍNH TIỀN - WORLD CUP 2026"
    ws["A1"].font      = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill      = HEADER
    ws["A1"].alignment = center

    total_cols = 2 + len(done_matches) + 3  # STT+Tên + các trận + Tổng nợ + Đã đóng + Còn lại
    if total_cols > 2:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws["A1"].fill = HEADER

    ws["A2"] = f"Xuất lúc: {datetime.now(TZ).strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(italic=True, size=10)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)

    # ---- Header hàng 3: STT | Tên | các trận | Tổng nợ | Đã đóng | Còn lại ----
    headers = ["STT", "Họ tên"]
    for mid, m in done_matches.items():
        ko  = datetime.fromisoformat(m["kickoff"]).strftime("%d/%m")
        fee = FEE.get(m["round"], 50000)
        hdr = f"{mid}\n{m['home_team']}\nvs\n{m['away_team']}\n{ko}\n(-{fee//1000}k)"
        headers.append(hdr)
    headers += ["Tổng nợ", "Đã đóng", "Còn lại"]

    for col, hdr in enumerate(headers, 1):
        cell            = ws.cell(row=3, column=col, value=hdr)
        cell.font       = bold_white
        cell.fill       = HEADER
        cell.alignment  = center
        cell.border     = thin
        ws.column_dimensions[get_column_letter(col)].width = 14 if col > 2 else (5 if col == 1 else 22)

    ws.row_dimensions[3].height = 80

    # ---- Dữ liệu từng thành viên ----
    for row_idx, (uid, member) in enumerate(members.items(), 4):
        player    = data["players"].get(uid, {})
        preds_uid = {mid: data["predictions"].get(mid, {}).get(uid) for mid in done_matches}

        ws.cell(row=row_idx, column=1, value=row_idx - 3).alignment = center
        name_cell            = ws.cell(row=row_idx, column=2, value=member["name"])
        name_cell.font       = bold_black
        name_cell.alignment  = Alignment(vertical="center")

        # Từng trận
        for col_idx, (mid, m) in enumerate(done_matches.items(), 3):
            fee    = FEE.get(m["round"], 50000)
            result = m.get("result")
            pred   = preds_uid.get(mid)
            cell   = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = center
            cell.border    = thin

            if pred is None:
                # Không bình chọn = thua
                cell.value = f"KBQ\n-{fee//1000}k"
                cell.fill  = YELLOW
                cell.font  = Font(color="7F6000")
            elif pred["choice"] == result:
                cell.value = "THẮNG"
                cell.fill  = GREEN
                cell.font  = Font(color="375623")
            else:
                choice_label = {"home": "TRÊN", "draw": "HÒA", "away": "DƯỚI"}.get(pred["choice"], pred["choice"])
                cell.value = f"THUA\n({choice_label})\n-{fee//1000}k"
                cell.fill  = RED
                cell.font  = Font(color="9C0006")

        # Tổng nợ / Đã đóng / Còn lại
        debt   = player.get("debt", 0)
        paid   = member.get("paid", 0)
        remain = max(0, debt - paid)
        total_col  = 3 + len(done_matches)
        paid_col   = total_col + 1
        remain_col = total_col + 2

        debt_cell          = ws.cell(row=row_idx, column=total_col,  value=debt)
        debt_cell.number_format = '#,##0"đ"'
        debt_cell.font     = bold_black
        debt_cell.fill     = RED if debt > 0 else GREEN
        debt_cell.alignment = center
        debt_cell.border   = thin

        paid_cell          = ws.cell(row=row_idx, column=paid_col,   value=paid)
        paid_cell.number_format = '#,##0"đ"'
        paid_cell.fill     = BLUE
        paid_cell.alignment = center
        paid_cell.border   = thin

        remain_cell        = ws.cell(row=row_idx, column=remain_col, value=remain)
        remain_cell.number_format = '#,##0"đ"'
        remain_cell.font   = bold_black
        remain_cell.fill   = RED if remain > 0 else GREEN
        remain_cell.alignment = center
        remain_cell.border = thin

        ws.row_dimensions[row_idx].height = 40

    # ---- Hàng tổng cộng ----
    last_row = 3 + len(members) + 1
    ws.cell(row=last_row, column=1, value="TỔNG").font = bold_black
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=2)

    total_col  = 3 + len(done_matches)
    paid_col   = total_col + 1
    remain_col = total_col + 2

    total_debt   = sum(data["players"].get(uid, {}).get("debt",  0) for uid in members)
    total_paid   = sum(members[uid].get("paid", 0)                  for uid in members)
    total_remain = sum(max(0, data["players"].get(uid, {}).get("debt", 0) - members[uid].get("paid", 0)) for uid in members)

    for col, val in [(total_col, total_debt), (paid_col, total_paid), (remain_col, total_remain)]:
        c = ws.cell(row=last_row, column=col, value=val)
        c.number_format = '#,##0"đ"'
        c.font  = bold_black
        c.fill  = GRAY
        c.alignment = center
        c.border = thin

    ws.cell(row=last_row, column=1).fill = GRAY
    ws.cell(row=last_row, column=2).fill = GRAY

    # ---- Chú thích màu ----
    note_row = last_row + 2
    ws.cell(row=note_row, column=1, value="CHÚ THÍCH:").font = bold_black
    notes = [
        (GREEN,  "THẮNG kèo"),
        (RED,    "THUA kèo"),
        (YELLOW, "KBQ = Không bình chọn (tính thua)"),
        (BLUE,   "Đã đóng tiền"),
    ]
    for i, (fill, text) in enumerate(notes):
        c = ws.cell(row=note_row + 1 + i, column=1, value=text)
        c.fill   = fill
        c.border = thin
        c.alignment = Alignment(vertical="center")
        ws.row_dimensions[note_row + 1 + i].height = 20

    # ---- Lưu file ----
    filename = f"TinhTien_WorldCup2026_{datetime.now(TZ).strftime('%d-%m-%Y')}.xlsx"
    filepath = os.path.join(os.getcwd(), filename)
    wb.save(filepath)

    # Gửi file vào nhóm Telegram
    try:
        with open(filepath, "rb") as f:
            await context.bot.send_document(
                chat_id=GROUP_ID,
                document=f,
                filename=filename,
                caption=f"BẢNG TÍNH TIỀN WORLD CUP 2026\nXuất lúc {datetime.now(TZ).strftime('%d/%m/%Y %H:%M')}\nTổng chưa thu: {total_remain:,}đ"
            )
        await update.message.reply_text(
            f"Đã xuất Excel và gửi vào nhóm!\n"
            f"File cũng được lưu tại: {filepath}"
        )
    except Exception as e:
        await update.message.reply_text(
            f"Đã lưu file tại: {filepath}\n"
            f"Lỗi gửi Telegram: {e}"
        )


def main():
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("newmatch",       cmd_newmatch))
    app.add_handler(CommandHandler("fetchodds",      cmd_fetchodds))
    app.add_handler(CommandHandler("autosetup",      cmd_autosetup))
    app.add_handler(CommandHandler("sendpoll",       cmd_sendpoll))
    app.add_handler(CommandHandler("lockpoll",       cmd_lockpoll))
    app.add_handler(CommandHandler("result",         cmd_result))
    app.add_handler(CommandHandler("fetchresult",    cmd_fetchresult))
    app.add_handler(CommandHandler("paid",           cmd_paid))
    app.add_handler(CommandHandler("unpaid",         cmd_unpaid))
    app.add_handler(CommandHandler("join",           cmd_join))
    app.add_handler(CommandHandler("members",        cmd_members))
    app.add_handler(CommandHandler("removemember",   cmd_removemember))
    app.add_handler(CommandHandler("addmember",      cmd_addmember))
    app.add_handler(CommandHandler("excel",          cmd_excel))
    app.add_handler(CommandHandler("standings",      cmd_standings))
    app.add_handler(CommandHandler("matches",        cmd_matches))
    app.add_handler(CommandHandler("mypredictions",  cmd_mypredictions))
    app.add_handler(CommandHandler("help",           cmd_help))
    app.add_handler(PollAnswerHandler(handle_poll_answer))
    logger.info("Bot đang chạy...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
