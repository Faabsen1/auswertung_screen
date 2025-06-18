import pygetwindow as gw
from screeninfo import get_monitors
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import mss
import numpy as np
import cv2
import pytesseract
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import json
import time

# Tesseract-Pfad anpassen!
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\f.morfeld\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789+-'

POSITION_FILE = "fensterkonfiguration.json"
EXCEL_FILE = "werte.xlsx"

def select_window_or_screen_from_list():
    # Fenster sammeln
    windows = [w for w in gw.getAllWindows() if w.title and w.width > 0 and w.height > 0]
    # Monitore sammeln
    monitors = get_monitors()
    # Auswahl-Liste bauen
    entries = [f"Monitor {i+1}: {m.width}x{m.height} @ ({m.x},{m.y})" for i, m in enumerate(monitors)]
    entries += [f"Fenster: {w.title} ({w.left},{w.top},{w.width},{w.height})" for w in windows]
    root = tk.Tk()
    root.title("Fenster oder Bildschirm auswählen")
    tk.Label(root, text="Bitte wähle ein Fenster oder einen Bildschirm aus:").pack(padx=10, pady=10)
    combo = ttk.Combobox(root, state="readonly", width=80)
    combo['values'] = entries
    combo.current(0)
    combo.pack(padx=10, pady=10)
    selected = {'index': 0}
    def on_ok():
        selected['index'] = combo.current()
        root.destroy()
    tk.Button(root, text="OK", command=on_ok).pack(pady=10)
    root.mainloop()
    idx = selected['index']
    if idx < len(monitors):
        m = monitors[idx]
        # Rückgabe: [x1, y1, x2, y2] für Monitor
        return [m.x, m.y, m.x + m.width, m.y + m.height]
    else:
        w = windows[idx - len(monitors)]
        # Rückgabe: [x1, y1, x2, y2] für Fenster
        return [w.left, w.top, w.left + w.width, w.top + w.height]

def frage_konfiguration():
    if os.path.exists(POSITION_FILE):
        root = tk.Tk()
        root.withdraw()
        antwort = messagebox.askyesno(
            "Konfiguration",
            "Haben sich Fensterposition oder Wertebereiche verändert?"
        )
        root.destroy()
        if not antwort:
            with open(POSITION_FILE, "r") as f:
                return json.load(f)
    fenster_position = select_window_or_screen_from_list()
    root = tk.Tk()
    root.withdraw()
    anzahl_werte = simpledialog.askinteger("Anzahl Werte", "Wie viele Werte sollen gelesen werden?", minvalue=1)
    trigger_area, werte_bereiche = select_trigger_and_value_areas(anzahl_werte, fenster_position)
    konfig = {
        "fenster_position": fenster_position,
        "anzahl_werte": anzahl_werte,
        "trigger_area": trigger_area,
        "werte_bereiche": werte_bereiche
    }
    with open(POSITION_FILE, "w") as f:
        json.dump(konfig, f)
    return konfig

def select_value_areas(anzahl_werte, fenster_position):
    img = screenshot_window_area(fenster_position)
    screenshot_img = img.copy()
    bereiche = []

    def mark_area(event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            mark_area.start = (x, y)
        elif event == cv2.EVENT_LBUTTONUP:
            end = (x, y)
            bereiche.append([mark_area.start[0], mark_area.start[1], end[0], end[1]])
            cv2.rectangle(screenshot_img, mark_area.start, end, (0, 255, 0), 2)
            cv2.imshow("Wertebereich markieren", screenshot_img)
    mark_area.start = (0, 0)

    cv2.namedWindow("Wertebereich markieren")
    cv2.setMouseCallback("Wertebereich markieren", mark_area)
    cv2.imshow("Wertebereich markieren", screenshot_img)
    print(f">> Markiere {anzahl_werte} Wertebereiche im Fenster, jeweils mit der Maus, dann drücke 'q'")
    while len(bereiche) < anzahl_werte:
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cv2.destroyWindow("Wertebereich markieren")
    return bereiche

def screenshot_window_area(area):
    time.sleep(0.2)  # Kurze Pause, um sicherzustellen, dass das Fenster bereit ist
    left, top, right, bottom = area
    width = right - left
    height = bottom - top
    with mss.mss() as sct:
        monitor = {"left": int(left), "top": int(top), "width": int(width), "height": int(height)}
        img = np.array(sct.grab(monitor))
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        return img

def prepare_excel(file, anzahl_werte):
    if not os.path.exists(file):
        wb = Workbook()
        ws = wb.active
        header = ["Zeitstempel"] + [f"Wert {i+1}" for i in range(anzahl_werte)]
        ws.append(header)
        wb.save(file)
    wb = load_workbook(file)
    ws = wb.active
    return wb, ws

def normalize_value(val):
    # Entfernt Leerzeichen und wandelt Unicode-Minus in ASCII-Minus um
    return val.replace('−', '-').replace('–', '-').replace(' ', '').strip()

def select_trigger_and_value_areas(anzahl_werte, fenster_position):
    img = screenshot_window_area(fenster_position)
    screenshot_img = img.copy()
    bereiche = []

    # Trigger-Bereich markieren
    cv2.namedWindow("Trigger-Bereich markieren")
    cv2.imshow("Trigger-Bereich markieren", screenshot_img)
    print(">> Markiere zuerst den Trigger-Bereich (z.B. Sekunden) und drücke 'q'")
    trigger = []

    def mark_trigger(event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            trigger.clear()
            trigger.append((x, y))
        elif event == cv2.EVENT_LBUTTONUP:
            trigger.append((x, y))
            cv2.rectangle(screenshot_img, trigger[0], trigger[1], (0, 0, 255), 2)
            cv2.imshow("Trigger-Bereich markieren", screenshot_img)
    cv2.setMouseCallback("Trigger-Bereich markieren", mark_trigger)
    while len(trigger) < 2:
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cv2.destroyWindow("Trigger-Bereich markieren")
    trigger_area = [trigger[0][0], trigger[0][1], trigger[1][0], trigger[1][1]]

    # Wertebereiche markieren
    wertebereiche = []
    cv2.namedWindow("Wertebereiche markieren")
    cv2.imshow("Wertebereiche markieren", screenshot_img)
    print(f">> Markiere jetzt {anzahl_werte} Wertebereiche und drücke jeweils 'q'")
    def mark_value(event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            mark_value.start = (x, y)
        elif event == cv2.EVENT_LBUTTONUP:
            end = (x, y)
            wertebereiche.append([mark_value.start[0], mark_value.start[1], end[0], end[1]])
            cv2.rectangle(screenshot_img, mark_value.start, end, (0, 255, 0), 2)
            cv2.imshow("Wertebereiche markieren", screenshot_img)
    mark_value.start = (0, 0)
    cv2.setMouseCallback("Wertebereiche markieren", mark_value)
    while len(wertebereiche) < anzahl_werte:
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cv2.destroyWindow("Wertebereiche markieren")
    return trigger_area, wertebereiche

if __name__ == "__main__":
    konfig = frage_konfiguration()
    fenster_position = konfig["fenster_position"]
    anzahl_werte = konfig["anzahl_werte"]
    trigger_area = konfig["trigger_area"]
    werte_bereiche = konfig["werte_bereiche"]

    wb, ws = prepare_excel(EXCEL_FILE, anzahl_werte)
    letzter_wert1 = None

    letzte_triggersekunde = None

    # Am Anfang des Skripts (vor dem Hauptloop):
    start_time = time.time()

    while True:
        img = screenshot_window_area(fenster_position)

        # Trigger auslesen
        trigger_img = img[trigger_area[1]:trigger_area[3], trigger_area[0]:trigger_area[2]]
        trigger_value = normalize_value(pytesseract.image_to_string(trigger_img, config=config))

        # Werte auslesen
        werte = []
        for (x1, y1, x2, y2) in werte_bereiche:
            wert_img = img[y1:y2, x1:x2]
            wert_raw = pytesseract.image_to_string(wert_img, config=config)
            werte.append(normalize_value(wert_raw))

        triggersekunde = werte[0]  # Der erste Bereich ist jetzt der Sekunden-Trigger

        if letzte_triggersekunde is None:
            letzte_triggersekunde = triggersekunde

        if triggersekunde != letzte_triggersekunde:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws.append([timestamp] + werte)
            wb.save(EXCEL_FILE)
            print(f"Werte gespeichert: {werte}")
            letzte_triggersekunde = triggersekunde

        # Im Hauptloop, direkt vor dem Speichern:
        laufzeit = time.time() - start_time
        print(f"Laufzeit: {laufzeit:.2f} Sekunden")

        time.sleep(0.2)

