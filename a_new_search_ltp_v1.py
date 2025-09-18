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


# ==============================
# --- Position Monitoring -----
# ==============================
# --- Set square-off time ---
SQUARE_OFF_HOUR = 15  # 3 PM
SQUARE_OFF_MINUTE = 29  # 3:20 PM
ADJUSTMENT_HOUR_LIMIT = 14 # 2 PM

def monitor_positions(wb, executed_orders, entry_total, symbol, expiry, lot_size):
    log_event("\n🔄 Starting continuous monitoring based on new strategy...")
    trade_terminal = wb.sheets["Trade_Terminal"]

    open_positions = list(executed_orders)

    # Fetch initial Q values to establish the first baseline for the 30% rule
    initial_q_total, _, _ = fetch_qk_values(wb, open_positions)
    if initial_q_total == 0:
        log_event("Warning: Initial Q values are zero. Falling back to search prices for initial adjustment baseline.", level="warning")
        initial_q_total = entry_total

    log_event(f"📈 Initial Combined Premium (Q Total) for adjustments: {initial_q_total:.2f}", level="info")

    try:
        while True:
            now = datetime.now()

            # 1. Check for end-of-day square-off
            if now.hour == SQUARE_OFF_HOUR and now.minute >= SQUARE_OFF_MINUTE:
                log_event(f"⏰ Auto square-off all open positions at {now.strftime('%H:%M:%S')}!", level="warning")
                for order in open_positions:
                    trade_terminal.range(f"T{order['row']}").value = "True_Market"
                    log_event(f"✍️ Marked row {order['row']} ({order['option_symbol']}) for Square_Off")
                break # Exit monitoring loop

            if not open_positions:
                log_event("No open positions to monitor. Exiting.", level="info")
                break

            # Fetch latest prices for all open positions
            q_total, k_total, order_qk = fetch_qk_values(wb, open_positions)

            # Log current status
            nifty_ltp = trade_terminal.range("K8").value
            log_event(f"📊 Status → Positions: {len(open_positions)} | Q Total: {q_total:.2f} | K Total: {k_total:.2f} | "
                      f"30% Trigger: {initial_q_total * 1.3:.2f} | NIFTY: {nifty_ltp}")

            # 2. Check for adjustment triggers (must be before 2 PM)
            adjustment_triggered = False
            if now.hour < ADJUSTMENT_HOUR_LIMIT:
                # Trigger A: Combined premium increased by 30%
                if k_total >= initial_q_total * 1.30:
                    log_event(f"🔥 Adjustment Trigger A: Combined premium {k_total:.2f} >= 30% threshold {initial_q_total * 1.30:.2f}", level="warning")
                    adjustment_triggered = True

                # Trigger B: One leg's premium dropped to 5-8 range
                if not adjustment_triggered:
                    for order in open_positions:
                        if 5 <= (order.get("k_value") or 0) <= 8:
                            log_event(f"🔥 Adjustment Trigger B: Leg {order['option_symbol']} premium is {order['k_value']:.2f} (in 5-8 range)", level="warning")
                            adjustment_triggered = True
                            break

            # 3. Perform Adjustment if triggered
            if adjustment_triggered:
                if len(open_positions) < 2:
                    log_event("⚠️ Adjustment triggered, but less than 2 open positions. Cannot perform adjustment.", level="warning")
                else:
                    # Identify profitable and unprofitable legs
                    profitable_leg = min(open_positions, key=lambda o: o.get("k_value") or float('inf'))
                    unprofitable_leg = max(open_positions, key=lambda o: o.get("k_value") or float('-inf'))

                    log_event(f"Identified Profitable Leg: {profitable_leg['option_symbol']} @ {profitable_leg['k_value']:.2f}", level="info")
                    log_event(f"Identified Unprofitable Leg: {unprofitable_leg['option_symbol']} @ {unprofitable_leg['k_value']:.2f}", level="info")

                    new_premium_target = unprofitable_leg.get("k_value")

                    # a. Close the profitable leg
                    log_event(f"Closing profitable leg: row {profitable_leg['row']}", level="info")
                    trade_terminal.range(f"T{profitable_leg['row']}").value = "True_Market"

                    # b. Remove from open positions list for monitoring
                    open_positions = [p for p in open_positions if p['row'] != profitable_leg['row']]

                    # c. Write the new leg
                    log_event(f"🔎 Searching for new {profitable_leg['option_type']} with premium near {new_premium_target:.2f}", level="info")
                    new_order = write_position(
                        wb, symbol, expiry, new_premium_target, lot_size,
                        option_type=profitable_leg['option_type'],
                        buy_or_sell="SELL"
                    )

                    if new_order:
                        log_event(f"✅ New position written: {new_order['option_symbol']}", level="info")
                        open_positions.append(new_order)

                        # d. Recalculate the baseline premium for the 30% rule
                        log_event("Waiting for new position's Q value to update...", level="info")
                        time.sleep(5) # Give Excel time to update Q value
                        new_q_total, _, _ = fetch_qk_values(wb, open_positions)
                        if new_q_total > 0:
                            initial_q_total = new_q_total
                            log_event(f"🔁 Baseline premium for 30% rule has been updated to: {initial_q_total:.2f}", level="info")
                        else:
                            log_event("⚠️ Could not get new Q values to update baseline. Baseline remains unchanged.", level="warning")

                    else:
                        log_event(f"⚠️ Could not find a new leg to write. Monitoring will continue with the remaining leg.", level="warning")

                    log_event("Pausing for 5 seconds after adjustment...", level="info")
                    time.sleep(5)
                    continue # Restart loop to get fresh data for all positions

            # Main loop sleep
            time.sleep(1)

    except Exception as e:
        log_event(f"❌ Critical error in monitoring: {str(e)}", level="error")
        import traceback
        log_event(f"Stack trace: {traceback.format_exc()}", level="error")
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
            log_event("Waiting for positions to initialize before monitoring...", level="info")
            time.sleep(5) # Give Excel time to populate Q values
            entry_total_fallback = (inputs["search_call_ltp"] or 0) + (inputs["search_put_ltp"] or 0)
            monitor_positions(wb, executed_orders, entry_total_fallback,
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
