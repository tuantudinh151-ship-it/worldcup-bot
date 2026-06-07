import os, json, logging, asyncio
from datetime import datetime
import pytz, requests
from telegram import Update
from telegram.ext import Application, CommandHandler, PollAnswerHandler, ContextTypes
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ============================================================
#  CẤU HÌNH - CHỈ SỬA PHẦN NÀY
# ============================================================
BOT_TOKEN   = "8831647645:AAGPwzT0zUu8eK7KULZXL1pDoRPKbUizgqU"
GROUP_ID    = -4992891193      # Chat ID nhóm (số âm)
ADMIN_ID    = 1216368366          # ID cá nhân của bạn
ODDS_API_KEY = "1c09f1cba53e1b6cc8d53d60f00501d5"
TIMEZONE    = "Asia/Ho_Chi_Minh"
# ============================================================

FEE = {"group": 50000, "knockout": 100000, "final": 200000}
SPORT_KEY = "soccer_fifa_world_cup"
TZ = pytz.timezone(TIMEZONE)
DATA_FILE = "data.json"
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Vòng đấu theo ngày (World Cup 2026)
# Vòng bảng: 11/06 - 01/07 | Vòng 1/16: 05/07 - 08/07
# Tứ kết: 11/07 - 12/07 | Bán kết: 15/07 - 16/07 | CK: 19/07
def get_round_type(commence_time: datetime) -> str:
    d = commence_time.astimezone(TZ).date()
    from datetime import date
    if d <= date(2026, 7, 1):  return "group"
    if d <= date(2026, 7, 12): return "knockout"
    return "final"

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
    """Lấy kèo từ The Odds API. market='spreads' (châu Á) hoặc 'h2h' (châu Âu)"""
    try:
        url = f"https://api.the-odds-api.com/v4/sports/{SPORT_KEY}/odds"
        params = {
            "apiKey": ODDS_API_KEY,
            "regions": "eu",
            "markets": market,
            "oddsFormat": "decimal",
            "dateFormat": "iso"
        }
        r = requests.get(url, params=params, timeout=15)
        if r.status_code != 200:
            logger.error(f"Odds API lỗi: {r.status_code} {r.text}")
            return []
        return r.json()
    except Exception as e:
        logger.error(f"fetch_odds lỗi: {e}")
        return []

def parse_asian_handicap(game: dict) -> dict | None:
    """Phân tích kèo châu Á từ 1 trận"""
    try:
        bookmakers_priority = ["pinnacle", "betfair_ex_eu", "betonlineag", "mybookieag"]
        chosen_bm = None

        # Ưu tiên nhà cái uy tín
        for bm_key in bookmakers_priority:
            for bm in game.get("bookmakers", []):
                if bm["key"] == bm_key:
                    chosen_bm = bm
                    break
            if chosen_bm:
                break

        # Nếu không tìm thấy nhà cái ưu tiên, lấy cái đầu tiên
        if not chosen_bm and game.get("bookmakers"):
            chosen_bm = game["bookmakers"][0]

        if not chosen_bm:
            return None

        market = next((m for m in chosen_bm["markets"] if m["key"] == "spreads"), None)
        if not market:
            return None

        outcomes = market["outcomes"]
        if len(outcomes) < 2:
            return None

        home = next((o for o in outcomes if o["name"] == game["home_team"]), outcomes[0])
        away = next((o for o in outcomes if o["name"] == game["away_team"]), outcomes[1])

        handicap = home.get("point", 0)  # số âm = đội nhà chấp, dương = đội nhà được chấp

        return {
            "home_team": game["home_team"],
            "away_team": game["away_team"],
            "commence_time": game["commence_time"],
            "match_id": game["id"][:8].upper(),
            "bookmaker": chosen_bm["title"],
            "handicap": handicap,       # mức chấp của đội nhà
            "home_odds": home["price"], # tỷ lệ cược đội nhà
            "away_odds": away["price"], # tỷ lệ cược đội khách
        }
    except Exception as e:
        logger.error(f"parse_asian_handicap lỗi: {e}")
        return None

def format_handicap(handicap: float, home_team: str, away_team: str) -> str:
    """Chuyển số handicap thành text dễ hiểu"""
    if handicap == 0:
        return "Kèo chẵn (0)"
    elif handicap < 0:
        # Đội nhà chấp
        h = abs(handicap)
        return f"{home_team} chấp {h} trái"
    else:
        # Đội nhà được chấp
        return f"{away_team} chấp {handicap} trái"

# ============================================================
#  LỆNH ADMIN
# ============================================================

async def cmd_fetchodds(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Lấy kèo mới nhất từ API và hiển thị danh sách trận. /fetchodds"""
    if not is_admin(update.effective_user.id):
        return

    await update.message.reply_text("Đang lấy kèo từ The Odds API...")

    games = fetch_odds("spreads")
    if not games:
        await update.message.reply_text("Không lấy được dữ liệu. Kiểm tra lại API key.")
        return

    now = datetime.now(TZ)
    upcoming = []
    for g in games:
        parsed = parse_asian_handicap(g)
        if not parsed:
            continue
        kickoff = datetime.fromisoformat(parsed["commence_time"].replace("Z", "+00:00")).astimezone(TZ)
        if kickoff > now:
            upcoming.append((kickoff, parsed))

    upcoming.sort(key=lambda x: x[0])

    if not upcoming:
        await update.message.reply_text("Không có trận nào sắp tới.")
        return

    msg = f"DANH SÁCH TRẬN SẮP TỚI ({len(upcoming)} trận)\n"
    msg += "Dùng /autosetup <số_trận> để tự động tạo poll\n"
    msg += "=" * 32 + "\n"

    for i, (kickoff, p) in enumerate(upcoming[:15], 1):
        hcap_text = format_handicap(p["handicap"], p["home_team"], p["away_team"])
        msg += (
            f"{i}. {p['home_team']} vs {p['away_team']}\n"
            f"   {kickoff.strftime('%d/%m %H:%M')} | {hcap_text}\n"
            f"   Tỷ lệ: Trên={p['home_odds']} / Dưới={p['away_odds']}\n"
            f"   ID tạm: {p['match_id']} | NHÀ CÁI: {p['bookmaker']}\n\n"
        )

    await update.message.reply_text(msg)
    context.user_data["upcoming_odds"] = upcoming


async def cmd_autosetup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Tự động tạo tất cả trận từ API và lên lịch gửi poll + khóa.
    /autosetup        → tạo tất cả trận chưa có
    /autosetup 5      → chỉ tạo 5 trận tiếp theo
    """
    if not is_admin(update.effective_user.id):
        return

    limit = int(context.args[0]) if context.args else 999
    await update.message.reply_text(f"Đang lấy kèo và tạo trận (tối đa {limit})...")

    games = fetch_odds("spreads")
    if not games:
        await update.message.reply_text("Không lấy được dữ liệu API.")
        return

    data = load_data()
    now = datetime.now(TZ)
    scheduler = context.bot_data.get("scheduler")
    created = 0
    skipped = 0

    for g in sorted(games, key=lambda x: x["commence_time"]):
        if created >= limit:
            break

        parsed = parse_asian_handicap(g)
        if not parsed:
            continue

        kickoff = datetime.fromisoformat(parsed["commence_time"].replace("Z", "+00:00")).astimezone(TZ)
        if kickoff <= now:
            continue

        match_id = parsed["match_id"]
        # Tránh trùng tên trận
        base_id = match_id
        counter = 1
        while match_id in data["matches"] and data["matches"][match_id]["home_team"] != parsed["home_team"]:
            match_id = f"{base_id}{counter}"
            counter += 1

        if match_id in data["matches"]:
            skipped += 1
            continue

        round_type = get_round_type(kickoff)
        fee = FEE[round_type]

        data["matches"][match_id] = {
            "home_team": parsed["home_team"],
            "away_team": parsed["away_team"],
            "kickoff": kickoff.isoformat(),
            "round": round_type,
            "handicap": parsed["handicap"],
            "home_odds": parsed["home_odds"],
            "away_odds": parsed["away_odds"],
            "bookmaker": parsed["bookmaker"],
            "result": None,
            "poll_id": None,
            "poll_message_id": None,
            "locked": False
        }

        # Lên lịch gửi poll trước kickoff 3 tiếng
        send_time = kickoff - __import__('datetime').timedelta(hours=3)
        if send_time > now and scheduler:
            scheduler.add_job(
                auto_send_poll_job, "date",
                run_date=send_time,
                args=[match_id, context.application],
                id=f"send_{match_id}", replace_existing=True
            )

        # Lên lịch khóa poll lúc kickoff
        if scheduler:
            scheduler.add_job(
                lock_poll_job, "date",
                run_date=kickoff,
                args=[match_id, context.application],
                id=f"lock_{match_id}", replace_existing=True
            )

        created += 1

    save_data(data)
    await update.message.reply_text(
        f"Đã tạo {created} trận mới, bỏ qua {skipped} trận đã có.\n"
        f"Poll sẽ tự động gửi vào nhóm trước 3 tiếng mỗi trận.\n"
        f"Dùng /matches để xem danh sách."
    )


async def auto_send_poll_job(match_id: str, app):
    """Tự động gửi poll vào nhóm trước kickoff"""
    data = load_data()
    if match_id not in data["matches"]:
        return
    m = data["matches"][match_id]
    if m.get("poll_message_id"):
        return  # Đã gửi rồi

    kickoff = datetime.fromisoformat(m["kickoff"])
    fee = FEE.get(m["round"], 50000)
    hcap_text = format_handicap(m["handicap"], m["home_team"], m["away_team"])

    question = (
        f"BÌNH CHỌN: {m['home_team']} vs {m['away_team']}\n"
        f"Kickoff: {kickoff.strftime('%d/%m/%Y %H:%M')} | Phí thua: {fee:,}đ\n"
        f"Kèo: {hcap_text}"
    )
    options = [
        f"TRÊN - {m['home_team']} ({m['home_odds']})",
        f"DƯỚI - {m['away_team']} ({m['away_odds']})"
    ]

    try:
        msg = await app.bot.send_poll(
            chat_id=GROUP_ID,
            question=question,
            options=options,
            is_anonymous=False,
            allows_multiple_answers=False
        )
        data["matches"][match_id]["poll_id"] = msg.poll.id
        data["matches"][match_id]["poll_message_id"] = msg.message_id
        save_data(data)
        logger.info(f"Đã tự động gửi poll trận {match_id}")
    except Exception as e:
        logger.error(f"Lỗi gửi poll {match_id}: {e}")


async def cmd_sendpoll(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Gửi poll thủ công. /sendpoll <mã_trận>"""
    if not is_admin(update.effective_user.id):
        return
    if not context.args:
        await update.message.reply_text("Cú pháp: /sendpoll <mã_trận>")
        return
    await auto_send_poll_job(context.args[0], context.application)
    await update.message.reply_text(f"Đã gửi poll trận {context.args[0]}.")


async def lock_poll_job(match_id: str, app):
    """Tự động khóa poll khi đến giờ kickoff"""
    data = load_data()
    if match_id not in data["matches"]:
        return
    m = data["matches"][match_id]
    if m["locked"] or not m.get("poll_message_id"):
        return
    try:
        await app.bot.stop_poll(chat_id=GROUP_ID, message_id=m["poll_message_id"])
        data["matches"][match_id]["locked"] = True
        save_data(data)
        await app.bot.send_message(
            chat_id=GROUP_ID,
            text=f"KHÓA BÌNH CHỌN\n{m['home_team']} vs {m['away_team']} - Bóng lăn rồi!"
        )
        logger.info(f"Đã khóa poll trận {match_id}")
    except Exception as e:
        logger.error(f"Lỗi khóa poll {match_id}: {e}")


async def cmd_result(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Nhập kết quả. /result <mã_trận> <home|away|draw>
    Ví dụ: /result ABC123 home   (đội nhà thắng/trên kèo)
            /result ABC123 away   (đội khách thắng/dưới kèo)
    Với kèo châu Á không có hòa — chỉ có home hoặc away.
    """
    if not is_admin(update.effective_user.id):
        return
    if len(context.args) < 2:
        await update.message.reply_text("Cú pháp: /result <mã_trận> <home|away>")
        return

    match_id = context.args[0].upper()
    result = context.args[1].lower()
    if result not in ["home", "away"]:
        await update.message.reply_text("Kết quả phải là: home hoặc away")
        return

    data = load_data()
    if match_id not in data["matches"]:
        await update.message.reply_text(f"Không tìm thấy trận {match_id}.")
        return

    m = data["matches"][match_id]
    data["matches"][match_id]["result"] = result
    fee = FEE.get(m["round"], 50000)
    hcap_text = format_handicap(m["handicap"], m["home_team"], m["away_team"])

    winners, losers = [], []
    for user_id, pred in data["predictions"].get(match_id, {}).items():
        player = data["players"].setdefault(user_id, {"name": pred["name"], "debt": 0})
        player["name"] = pred["name"]
        if pred["choice"] == result:
            winners.append(pred["name"])
        else:
            losers.append(pred["name"])
            player["debt"] = player.get("debt", 0) + fee

    save_data(data)

    result_label = m["home_team"] if result == "home" else m["away_team"]
    msg = (
        f"KẾT QUẢ TRẬN {match_id}\n"
        f"{m['home_team']} vs {m['away_team']}\n"
        f"Kèo: {hcap_text}\n"
        f"Kết quả kèo: {result_label} thắng kèo\n\n"
    )
    if winners:
        msg += f"THẮNG ({len(winners)} người): {', '.join(winners)}\n"
    if losers:
        msg += f"THUA ({len(losers)} người, -{fee:,}đ): {', '.join(losers)}\n"
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
    msg = "BẢNG XẾP HẠNG NỢ\n" + "=" * 28 + "\n"
    for i, (uid, p) in enumerate(sorted_p, 1):
        debt = p.get("debt", 0)
        emoji = "🔴" if debt > 0 else "🟢"
        msg += f"{i}. {emoji} {p['name']}: {'-' if debt > 0 else ''}{debt:,}đ\n"
    await update.message.reply_text(msg)


async def cmd_matches(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    now = datetime.now(TZ)
    upcoming = [(mid, m, datetime.fromisoformat(m["kickoff"]))
                for mid, m in data["matches"].items()
                if datetime.fromisoformat(m["kickoff"]) > now and not m.get("result")]
    upcoming.sort(key=lambda x: x[2])
    if not upcoming:
        await update.message.reply_text("Không có trận nào sắp tới.")
        return
    msg = f"CÁC TRẬN SẮP TỚI ({len(upcoming)} trận)\n" + "=" * 28 + "\n"
    for mid, m, kickoff in upcoming[:10]:
        status = "Đã gửi poll" if m.get("poll_message_id") else "Chưa gửi"
        hcap = format_handicap(m.get("handicap", 0), m["home_team"], m["away_team"])
        msg += f"{mid}: {m['home_team']} vs {m['away_team']}\n"
        msg += f"   {kickoff.strftime('%d/%m %H:%M')} | {hcap} | {status}\n"
    await update.message.reply_text(msg)


async def cmd_mypredictions(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = str(update.effective_user.id)
    data = load_data()
    history = []
    for mid, preds in data["predictions"].items():
        if user_id in preds:
            m = data["matches"].get(mid, {})
            result = m.get("result")
            pred = preds[user_id]["choice"]
            outcome = ("THẮNG" if pred == result else "THUA") if result else "Chờ kết quả"
            team = m.get("home_team","?") if pred == "home" else m.get("away_team","?")
            history.append(f"{mid}: Chọn {team} → {outcome}")
    if not history:
        await update.message.reply_text("Bạn chưa bình chọn trận nào.")
        return
    debt = data["players"].get(user_id, {}).get("debt", 0)
    msg = "LỊCH SỬ BÌNH CHỌN\n" + "\n".join(history) + f"\n\nTổng nợ: {debt:,}đ"
    await update.message.reply_text(msg)


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "HƯỚNG DẪN BOT WORLD CUP 2026\n"
        "=" * 30 + "\n\n"
        "LỆNH TẤT CẢ:\n"
        "/matches — Trận sắp tới\n"
        "/standings — Bảng nợ tiền\n"
        "/mypredictions — Lịch sử của bạn\n\n"
        "LỆNH ADMIN:\n"
        "/fetchodds — Xem kèo mới từ nhà cái\n"
        "/autosetup [số] — Tự động tạo trận + lên lịch\n"
        "/sendpoll <id> — Gửi poll thủ công\n"
        "/result <id> <home|away> — Nhập kết quả\n\n"
        "PHÍ THUA:\n"
        "Vòng bảng: 50,000đ\n"
        "Vòng loại trực tiếp: 100,000đ\n"
        "Chung kết: 200,000đ\n\n"
        "CÁCH CHƠI KÈO CHÂU Á:\n"
        "TRÊN = đội chấp thắng kèo\n"
        "DƯỚI = đội được chấp thắng kèo"
    )


# ============================================================
#  XỬ LÝ BÌNH CHỌN
# ============================================================

async def handle_poll_answer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = update.poll_answer
    poll_id = answer.poll_id
    user = answer.user
    user_id = str(user.id)

    if not answer.option_ids:
        return

    option_index = answer.option_ids[0]
    choice = "home" if option_index == 0 else "away"

    data = load_data()
    match_id = next((mid for mid, m in data["matches"].items() if m.get("poll_id") == poll_id), None)
    if not match_id or data["matches"][match_id]["locked"]:
        return

    data["predictions"].setdefault(match_id, {})[user_id] = {
        "name": user.full_name, "choice": choice
    }
    data["players"].setdefault(user_id, {"name": user.full_name, "debt": 0})["name"] = user.full_name
    save_data(data)


# ============================================================
#  KHỞI ĐỘNG
# ============================================================

async def post_init(application):
    scheduler = AsyncIOScheduler(timezone=TZ)
    scheduler.start()
    application.bot_data["scheduler"] = scheduler

    # Khôi phục lịch khóa poll
    data = load_data()
    now = datetime.now(TZ)
    for match_id, m in data["matches"].items():
        kickoff = datetime.fromisoformat(m["kickoff"])
        if not m["locked"] and m.get("poll_message_id") and kickoff > now:
            scheduler.add_job(lock_poll_job, "date", run_date=kickoff,
                args=[match_id, application], id=f"lock_{match_id}", replace_existing=True)

        # Khôi phục lịch gửi poll tự động
        if not m.get("poll_message_id") and kickoff > now:
            from datetime import timedelta
            send_time = kickoff - timedelta(hours=3)
            if send_time > now:
                scheduler.add_job(auto_send_poll_job, "date", run_date=send_time,
                    args=[match_id, application], id=f"send_{match_id}", replace_existing=True)

    logger.info("Bot đã khởi động!")


def main():
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("fetchodds", cmd_fetchodds))
    app.add_handler(CommandHandler("autosetup", cmd_autosetup))
    app.add_handler(CommandHandler("sendpoll", cmd_sendpoll))
    app.add_handler(CommandHandler("result", cmd_result))
    app.add_handler(CommandHandler("standings", cmd_standings))
    app.add_handler(CommandHandler("matches", cmd_matches))
    app.add_handler(CommandHandler("mypredictions", cmd_mypredictions))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(PollAnswerHandler(handle_poll_answer))
    logger.info("Bot đang chạy...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()