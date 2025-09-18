import os
import sys
import warnings
import time
import logging
import os
from datetime import datetime
import signal

warnings.filterwarnings("ignore")

# --- Logging & Monitoring ----
# ==============================

# Create folder if not exists
log_dir = "Logs/personal_trade"
os.makedirs(log_dir, exist_ok=True)

# Generate filename with datetime
log_filename = datetime.now().strftime("trade_monitor_%Y%m%d_%H%M%S.log")
log_path = os.path.join(log_dir, log_filename)

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log_event(message, level="info"):
    """Log message to console + file"""
    print(message)
    if level == "info":
        logging.info(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "error":
        logging.error(message)

# Log program start
log_event("🚀 Program starting...", level="info")

# ==============================
# --- Dependency Checks -------
# ==============================
try:
    import xlwings as xw
except (ModuleNotFoundError, ImportError) as e:
    log_event(f"Installing xlwings... Error: {str(e)}", level="warning")
    os.system(f"{sys.executable} -m pip install -U xlwings")
    import xlwings as xw

try:
    import pandas as pd
except (ModuleNotFoundError, ImportError) as e:
    log_event(f"Installing pandas... Error: {str(e)}", level="warning")
    os.system(f"{sys.executable} -m pip install -U pandas")
    import pandas as pd


# ==============================
# --- Workbook Setup ----------
# ==============================
def setup_workbook(filename="Finvasia_Trade_Terminal_v3.xlsm"):
    if not os.path.exists(filename):
        wb = xw.Book()
        wb.save(filename)
        wb.close()

    wb = xw.Book(filename)
    required_sheets = ["Trade_Terminal", "Option_Chain_Input", "Option_Chain_Output", "Chartink_Result"]
    for sheet in required_sheets:
        if sheet not in [s.name for s in wb.sheets]:
            wb.sheets.add(sheet)

    return wb


# ==============================
# --- Excel Access Helpers ----
# ==============================
def read_inputs(wb):
    trade_terminal = wb.sheets["Trade_Terminal"]
    option_chain_input = wb.sheets["Option_Chain_Input"]

    symbol_name = option_chain_input.range("E3").value
    date_value = option_chain_input.range("E4").value
    lot_size = int(option_chain_input.range("C2").value or 1)

    search_call_ltp = trade_terminal.range("AH2").value
    search_put_ltp = trade_terminal.range("AI2").value

    final_expiry_date = None
    if date_value:
        final_expiry_date = datetime.strptime(str(date_value), "%Y-%m-%d %H:%M:%S").strftime("%d%b%y").upper()

    return {
        "symbol": symbol_name,
        "expiry": final_expiry_date,
        "lot_size": lot_size,
        "search_call_ltp": search_call_ltp,
        "search_put_ltp": search_put_ltp,
    }


# ==============================
# --- Position Writer ---------
# ==============================
def write_position(wb, symbol, expiry, search_ltp, lot_size,
                   option_type="CALL", buy_or_sell="SELL", entry_signal="True_Market"):

    trade_terminal = wb.sheets["Trade_Terminal"]
    option_chain_output = wb.sheets["Option_Chain_Output"]

    def find_nearest_row(sheet, column_letter, search_value):
        values = sheet.range(f"{column_letter}:{column_letter}").value
        numeric_values = [(i + 1, float(v)) for i, v in enumerate(values) if isinstance(v, (int, float))]
        if not numeric_values:
            return None
        return min(numeric_values, key=lambda x: abs(x[1] - float(search_value)))

    def find_last_row(sheet, col="A"):
        return sheet.range(col + str(sheet.cells.last_cell.row)).end("up").row + 1

    # Column mapping based on option type
    if option_type.upper() == "CALL":
        strike_col = "J"
        strike_read_col = "P"
    elif option_type.upper() == "PUT":
        strike_col = "V"
        strike_read_col = "P"
    else:
        print(f"⚠️ Invalid option_type: {option_type}")
        return None

    opt_row = find_nearest_row(option_chain_output, strike_col, search_ltp)
    if not opt_row:
        # Instead of:
        print(f"⚠️ No valid strike found in Option_Chain_Output for {option_type}.")
        
        # Use:
        log_event(f"⚠️ No valid strike found in Option_Chain_Output for {option_type}.", level="warning")
        return None

    strike_cell_value = option_chain_output.range(f"{strike_read_col}{opt_row[0]}").value
    if strike_cell_value is None:
        print(f"⚠️ Strike price cell is empty at {strike_read_col}{opt_row[0]} for {option_type}")
        return None

    strike_price = int(strike_cell_value)

    option_symbol = f"NFO:{symbol}{expiry}{'C' if option_type.upper()=='CALL' else 'P'}{strike_price}"

    target_row = find_last_row(trade_terminal, "A")

    # Write to Excel
    trade_terminal.range(f"A{target_row}").value = option_symbol
    trade_terminal.range(f"M{target_row}").value = lot_size
    trade_terminal.range(f"N{target_row}").value = buy_or_sell
    trade_terminal.range(f"O{target_row}").value = entry_signal

    print(f"✅ {option_type} written → {option_symbol} at row {target_row} (Q pending)")

    return {
        "row": target_row,
        "option_symbol": option_symbol,
        "lot_size": lot_size,
        "buy_or_sell": buy_or_sell,
        "option_type": option_type,
        "q_value": None,
        "k_value": None,
    }


# ==============================
# --- Fetch Q & K Values ------
# ==============================
# ==============================
# --- Fetch Q & K Values ------
# ==============================
def fetch_qk_values(wb, executed_orders, per_order_timeout=5.0):
    """
    Read Q and K values for all executed_orders.
    Waits up to per_order_timeout seconds for each order's Q cell to become non-None.
    Returns q_total, k_total, order_qk (dict mapping row -> (q,k)).
    """
    trade_terminal = wb.sheets["Trade_Terminal"]
    order_qk = {}
    q_total, k_total = 0.0, 0.0

    for order in executed_orders:
        if not order:
            continue

        row = order["row"]
        q_value = None
        start = time.time()
        # wait up to per_order_timeout seconds for Q cell to populate
        while q_value is None and (time.time() - start) < per_order_timeout:
            try:
                q_val = trade_terminal.range(f"Q{row}").value
            except Exception:
                q_val = None
            if q_val is None:
                time.sleep(0.2)
            else:
                # ensure numeric
                try:
                    q_value = float(q_val)
                except Exception:
                    q_value = q_val
        # K value: read once (use 0 if missing)
        try:
            k_read = trade_terminal.range(f"K{row}").value
            k_value = float(k_read) if k_read is not None else 0.0
        except Exception:
            k_value = 0.0

        # keep order dict fields up-to-date
        order["q_value"] = q_value or 0.0
        order["k_value"] = k_value or 0.0

        q_total += order["q_value"]
        k_total += order["k_value"]
        order_qk[row] = (order["q_value"], order["k_value"])

    return q_total, k_total, order_qk


    # --- inside monitor_positions, replace the existing stoploss section with this ---
    # --- Stoploss adjustment ---
    # ===========================
    # Stoploss adjustment block
    # ===========================
    if k_total > stoploss_value:
        adjustment_count += 1
        log_event(
            f"⚠️ Adjustment {adjustment_count} Triggered! LTP {k_total} > StopLoss {stoploss_value:.2f}",
            level="warning"
        )

        # find the leg with the largest K (losing leg)
        losing_leg = max(executed_orders, key=lambda o: (o.get("k_value") or 0.0))
        losing_row = losing_leg["row"]

        current_t_flag = trade_terminal.range(f"T{losing_row}").value
        if str(current_t_flag).strip().lower() != "true_market":
            # 1. Close losing leg
            trade_terminal.range(f"T{losing_row}").value = "True_Market"
            log_event(
                f"✍️ Closed losing leg row {losing_row} ({losing_leg['option_symbol']}) at K={losing_leg['k_value']}",
                level="info"
            )

            # 2. Place opposite order
            new_order = search_opposite_and_write(wb, symbol, expiry, losing_leg, lot_size)
            if new_order:
                executed_orders.append(new_order)
                log_event(f"✅ Replacement {new_order['option_type']} placed at row {new_order['row']}", level="info")

                # 3. Refresh totals
                time.sleep(1)  # give Excel time to update
                q_total, k_total, order_qk = fetch_qk_values(wb, executed_orders)

                # 4. Update stoploss (dynamic)
                stoploss_value = q_total * 1.001
                log_event(
                    f"📊 After replacement → Entry Total (Q): {q_total} | Live Total (K): {k_total} | "
                    f"🔁 New StopLoss: {stoploss_value:.2f}",
                    level="info"
                )
            else:
                log_event("⚠️ Replacement order failed (no strike found).", level="warning")

        # cooldown → prevents multiple triggers on the *same tick*
        time.sleep(2)


# ===========================
# 2. Extra early-exit condition (1 < K < 5 and before 2 PM)
# ===========================
    elif any(1 <= o.get("k_value", 0) <= 5 for o in executed_orders) and now.hour < 14:
        target_leg = next(o for o in executed_orders if 1 <= o.get("k_value", 0) <= 5)

        log_event(f"⚠️ Early Exit Triggered → {target_leg['option_symbol']} has K={target_leg['k_value']} (<5) before 2PM", level="warning")

        losing_row = target_leg["row"]
        current_t_flag = trade_terminal.range(f"T{losing_row}").value
        if str(current_t_flag).strip().lower() != "true_market":
            trade_terminal.range(f"T{losing_row}").value = "True_Market"
            log_event(f"✍️ Closed leg row {losing_row} ({target_leg['option_symbol']}) at K={target_leg['k_value']}", level="info")

            # Place new opposite order
            new_order = search_opposite_and_write(wb, symbol, expiry, target_leg, lot_size)
            if new_order:
                executed_orders.append(new_order)

            # Recompute stoploss after replacement
            q_total, k_total, order_qk = fetch_qk_values(wb, executed_orders)
            # stoploss_value = q_total * 1.3
            stoploss_value = q_total * 1.001
            log_event(f"🔁 New StopLoss value set to {stoploss_value:.2f}", level="info")

        time.sleep(2)
# ==============================
# --- Search Opposite & Write --
# ==============================
def search_opposite_and_write(wb, symbol, expiry, losing_leg, lot_size):
    losing_type = losing_leg["option_type"]
    losing_ltp = losing_leg["k_value"]

    opposite_type = "PUT" if losing_type == "CALL" else "CALL"
    print(f"🔎 Searching {losing_ltp} in {opposite_type} Option Chain...")

    new_order = write_position(
        wb, symbol, expiry, losing_ltp, lot_size,
        option_type=opposite_type,
        buy_or_sell="SELL",
        entry_signal="True_Market"
    )

    if new_order:
        print(f"✅ New {opposite_type} SELL executed at row {new_order['row']}")
    return new_order


# ==============================
# --- Logging & Monitoring ----
# ==============================

# Create folder if not exists
log_dir = "Logs/personal_trade"
os.makedirs(log_dir, exist_ok=True)

# Generate filename with datetime
log_filename = datetime.now().strftime("trade_monitor_%Y%m%d_%H%M%S.log")
log_path = os.path.join(log_dir, log_filename)

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


def log_event(message, level="info"):
    """Log message to console + file"""
    print(message)
    if level == "info":
        logging.info(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "error":
        logging.error(message)


# --- Set square-off time ---
SQUARE_OFF_HOUR = 15  # 3 PM
SQUARE_OFF_MINUTE = 29  # 3:20 PM

def monitor_positions(wb, executed_orders, entry_total, symbol, expiry, lot_size):
    log_event("\n🔄 Starting continuous monitoring...")
    trade_terminal = wb.sheets["Trade_Terminal"]
    adjustment_count = 0

    try:
        nifty_ltp = trade_terminal.range("K8").value
        q_total, k_total, order_qk = fetch_qk_values(wb, executed_orders)

        # Initialize stoploss with first q_total
        # stoploss_value = q_total * 1.3
        stoploss_value = q_total * 1.001

        log_event(f"📊 Initial Trade → Entry Total (Q): {q_total} | Live Total (K): {k_total:.2f} | "
                  f"StopLoss: {stoploss_value:.2f} | NIFTY LTP: {nifty_ltp}")

        while True:
            now = datetime.now()
            if now.hour == SQUARE_OFF_HOUR and now.minute >= SQUARE_OFF_MINUTE:
                log_event(f"⏰ Auto square-off all positions!", level="warning")
                for order in executed_orders:
                    row = order["row"]
                    trade_terminal.range(f"T{row}").value = "True_Market"
                    log_event(f"✍️ Marked row {row} ({order['option_symbol']}) as Square_Off")
                break

            nifty_ltp = trade_terminal.range("K8").value

            # --- get latest totals ---
            q_total, k_total, order_qk = fetch_qk_values(wb, executed_orders)

            # 🔁 Always recompute stoploss dynamically
            # stoploss_value = q_total * 1.3
            stoploss_value = q_total * 1.001

            log_event(f"📊 Adjustment {adjustment_count} → Entry Total (Q): {q_total} | Live Total (K): {k_total} | "
                      f"StopLoss: {stoploss_value:.2f} | NIFTY LTP: {nifty_ltp}")

            # ===========================
            # Stoploss adjustment block
            # ===========================
            if k_total > stoploss_value:
                adjustment_count += 1
                log_event(f"⚠️ Adjustment {adjustment_count} Triggered! LTP {k_total} > StopLoss {stoploss_value:.2f}", level="warning")

                # (close losing leg + replacement order code here)
                # ...

            time.sleep(1)

    except Exception as e:
        log_event(f"❌ Critical error in monitoring: {str(e)}", level="error")
    finally:
        log_event("🏁 Monitoring completed", level="info")




# ==============================
# --- Main --------------------
# ==============================
if __name__ == "__main__":
    try:
        log_event("📘 Starting program execution")
        wb = setup_workbook()
        inputs = read_inputs(wb)

        log_event("📘 Workbook and sheets ready.")
        log_event("➡️ Input values loaded:")
        for k, v in inputs.items():
            log_event(f"   {k}: {v}")

        executed_orders = []

        # Execute CALL
        call_order = write_position(
            wb, inputs["symbol"], inputs["expiry"],
            inputs["search_call_ltp"], inputs["lot_size"],
            option_type="CALL", buy_or_sell="SELL"
        )
        if call_order:
            executed_orders.append(call_order)

        # Execute PUT
        put_order = write_position(
            wb, inputs["symbol"], inputs["expiry"],
            inputs["search_put_ltp"], inputs["lot_size"],
            option_type="PUT", buy_or_sell="SELL"
        )
        if put_order:
            executed_orders.append(put_order)

        if not executed_orders:
            log_event("❌ No positions executed. Exiting.", level="warning")
        else:
            entry_total = inputs["search_call_ltp"] + inputs["search_put_ltp"]
            monitor_positions(wb, executed_orders, entry_total,
                            inputs["symbol"], inputs["expiry"], inputs["lot_size"])
    except Exception as e:
        log_event(f"❌ CRITICAL ERROR: {str(e)}", level="error")
        import traceback
        log_event(f"Stack trace: {traceback.format_exc()}", level="error")
    finally:
        log_event("🏁 Program execution completed", level="info")





def signal_handler(sig, frame):
    log_event("⚠️ Program interrupted by user", level="warning")
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)
