import base64
import json
import os
import threading
import time
import tkinter as tk
from ctypes import cast, POINTER
from io import BytesIO
from tkinter import ttk

import numpy as np
import pyaudio
import pygetwindow as gw
from PIL import Image
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pynput import keyboard
from pynput.keyboard import Controller
from pystray import Icon, MenuItem as item


# ==================== ФУНКЦИИ РАБОТЫ С КЛАВИШАМИ ====================

def str_to_keys(keys_str):
    """
    Преобразует строку (например, "ctrl + t") в список объектов клавиш pynput.
    """
    keys = []
    for key_str in keys_str.split(' + '):
        key_str = key_str.strip()
        try:
            # Для специальных клавиш (ctrl, alt, shift и т.п.)
            key = getattr(keyboard.Key, key_str)
            keys.append(key)
        except AttributeError:
            # Для символьных клавиш (буквы, цифры)
            keys.append(keyboard.KeyCode.from_char(key_str))
    return keys


def keys_to_str(keys):
    """
    Преобразует список объектов клавиш в строковое представление.
    """
    non_char_keys = []
    char_keys = []
    for key in keys:
        if hasattr(key, 'char') and key.char is not None:
            char_keys.append(key.char)  # Убираем .lower()
        else:
            non_char_keys.append(str(key).replace("Key.", ""))
    return ' + '.join(non_char_keys + char_keys)


def get_pressed_keys(entry_field):
    """
    Запускает временный listener, ожидающий нажатия клавиши,
    затем обновляет содержимое текстового поля (entry_field) выбранной комбинацией.
    """
    keys_pressed = []
    pressed_codes = []

    def on_press(key):
        if key not in pressed_codes:
            pressed_codes.append(key)
            if hasattr(key, 'char') and key.char is not None:
                keys_pressed.append(keyboard.KeyCode.from_char(key.char.lower().strip("'")))
            else:
                keys_pressed.append(key)
            # После первого нажатия завершаем listener
            return False

    def on_release(key):
        return False

    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()

    keys_str = keys_to_str(keys_pressed)
    entry_field.delete(0, tk.END)
    entry_field.insert(0, keys_str)
    return keys_str


# ==================== ЗАГРУЗКА И СОХРАНЕНИЕ НАСТРОЕК ====================

def load_settings():
    """
    Загружает настройки из файла settings.json или создаёт его со значениями по умолчанию.
    """
    settings_file = "talk-to-press-settings.json"
    default_settings = {
        "volume_threshold": 700,  # Порог громкости для активации PTT
        "ptt_keys_str": "t",  # Клавиша push-to-talk (можно задать комбинацию, например, "ctrl + t")
        "allowed_window_fragments": "squad, company",  # Фрагменты названия окна, при наличии которых PTT активен
        "post_voice_release_delay": 800,  # Задержка (мс) перед отпусканием клавиш после окончания речи
        "microphone_index": 0,  # Индекс микрофонного устройства
        "ignore_keys_enabled": False,  # Флаг: игнорировать PTT, если нажаты определённые клавиши
        "ignore_keys_str": "v + b",  # Строка с клавишами для игнорирования PTT
        "fade_sound_enabled": False,  # Флаг: затемнять звук динамиков во время разговора
        "fade_sound_percentage": 90,  # Процент уменьшения громкости при активации PTT
        "mute_all_enabled": False,  # Флаг: включать режим mute для динамиков
        "mute_key": "m"  # Клавиша для режима mute
    }
    if os.path.exists(settings_file):
        with open(settings_file, 'r') as file:
            try:
                return json.load(file)
            except json.JSONDecodeError:
                return default_settings
    else:
        with open(settings_file, 'w') as file:
            json.dump(default_settings, file, indent=4)
        return default_settings


def save_settings():
    """
    Сохраняет настройки, введённые в окне настроек, в файл settings.json и применяет их.
    """
    global volume_threshold, ptt_keys_str, allowed_window_fragments, post_voice_release_delay, \
        selected_mic_index, ignore_keys_enabled, ignore_keys_str, fade_sound_enabled, fade_sound_percentage, \
        mute_all_enabled, mute_key, ignore_key_codes, mute_key_codes, ptt_key_codes

    # Получаем выбранный микрофон по имени из выпадающего списка
    new_mic_index = microphones.index(next(mic for mic in microphones if mic[1] == mic_choice_var.get()))
    settings = {
        "volume_threshold": int(volume_threshold_scale.get()),
        "ptt_keys_str": ptt_key_entry.get(),
        "allowed_window_fragments": window_entry.get().lower(),
        "post_voice_release_delay": int(delay_entry.get()),
        "microphone_index": new_mic_index,
        "ignore_keys_enabled": ignore_keys_checkbox.get(),
        "ignore_keys_str": ignore_keys_entry.get(),
        "fade_sound_enabled": fade_sound_checkbox_var.get(),
        "fade_sound_percentage": int(fade_sound_percent_combobox.get()),
        "mute_all_enabled": mute_all_checkbox.get(),
        "mute_key": mute_key_entry.get()
    }
    with open("talk-to-press-settings.json", 'w') as file:
        json.dump(settings, file, indent=4)

    # Применяем новые настройки
    volume_threshold = settings["volume_threshold"]
    ptt_keys_str = settings["ptt_keys_str"]
    allowed_window_fragments = settings["allowed_window_fragments"]
    post_voice_release_delay = settings["post_voice_release_delay"]
    ignore_keys_enabled = settings["ignore_keys_enabled"]
    ignore_keys_str = settings["ignore_keys_str"]
    fade_sound_enabled = settings["fade_sound_enabled"]
    fade_sound_percentage = settings["fade_sound_percentage"]
    mute_all_enabled = settings["mute_all_enabled"]
    mute_key = settings["mute_key"]

    if new_mic_index != selected_mic_index:
        selected_mic_index = new_mic_index
        set_microphone_device(microphones[selected_mic_index][0])
    # Пересоздаем игнорируемые клавиши после обновления настроек
    ignore_key_codes = str_to_keys(ignore_keys_str)
    ptt_key_codes = str_to_keys(ptt_keys_str)
    mute_key_codes = str_to_keys(mute_key)


# Загружаем настройки при запуске
settings = load_settings()
volume_threshold = settings["volume_threshold"]
ptt_keys_str = settings["ptt_keys_str"]
allowed_window_fragments = settings["allowed_window_fragments"]
post_voice_release_delay = settings["post_voice_release_delay"]
selected_mic_index = settings["microphone_index"]
ignore_keys_enabled = settings["ignore_keys_enabled"]
ignore_keys_str = settings["ignore_keys_str"]
fade_sound_enabled = settings["fade_sound_enabled"]
fade_sound_percentage = settings["fade_sound_percentage"]
mute_all_enabled = settings["mute_all_enabled"]
mute_key = settings["mute_key"]
# ==================== АУДИО НАСТРОЙКИ ====================
chunk = 2205
audio_format = pyaudio.paInt16
channels = 1
rate = 22050

# Контроллер для эмуляции нажатия клавиш
keyboard_controller = Controller()
p = pyaudio.PyAudio()


def get_available_microphones():
    """
    Возвращает список доступных микрофонных устройств в формате [(индекс, имя), ...].
    """
    return [
        (p.get_device_info_by_index(i)['index'], p.get_device_info_by_index(i)['name'])
        for i in range(p.get_device_count())
        if p.get_device_info_by_index(i)['maxInputChannels'] > 0
    ]


microphones = get_available_microphones()

# ==================== TKINTER И ОКНО НАСТРОЕК ====================
root = tk.Tk()
root.withdraw()  # Скрываем главное окно

volume_var = tk.DoubleVar()


def get_active_window():
    """
    Возвращает название активного окна (в нижнем регистре).
    """
    try:
        return gw.getActiveWindow().title.lower()
    except AttributeError:
        return ""


# Глобальные переменные для состояния push-to-talk
is_talking = False
last_above_threshold_time = 0
original_speaker_volume = None  # Для восстановления громкости после затемнения
stored_speaker_volume = None  # Для восстановления громкости после режима mute
original_microphone_volume = None
stored_microphone_volume = None

# Получаем устройство динамиков для управления громкостью
speakers = AudioUtilities.GetSpeakers()
interface = speakers.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume_control = cast(interface, POINTER(IAudioEndpointVolume))

# Получаем устройство микрофона для управления громкостью

microphone_devices = AudioUtilities.GetMicrophone()
interface = microphone_devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
microphone_volume = cast(interface, POINTER(IAudioEndpointVolume))

# ==================== ГЛОБАЛЬНЫЙ LISTENER ДЛЯ КЛАВИАТУРЫ ====================
pressed_keys_global = set()


def on_press_global(key):
    if key not in pressed_keys_global:
        pressed_keys_global.add(key)
        print(pressed_keys_global)


def on_release_global(key):
    if key in pressed_keys_global:
        time.sleep(0.01)
        pressed_keys_global.remove(key)


keyboard_listener = keyboard.Listener(on_press=on_press_global, on_release=on_release_global)
keyboard_listener.start()

muted = False  # Флаг состояния mute для динамиков


# ==================== ЛОКАЛЬНЫЙ LISTENER ДЛЯ КЛАВИАТУРЫ ====================
# Получение нажатых кнопок
def get_local_pressed_keys(field):
    local_keys_code_pressed = []  # Список для хранения нажатых клавиш
    local_keys_pressed = []  # Список для хранения нажатых клавиш

    def local_on_press(key):
        # Это чтобы одна клавиша не нажималась много раз
        if key not in local_keys_code_pressed:
            local_keys_code_pressed.append(key)
            # Если это символ (буква или цифра)
            if hasattr(key, 'char') and key.char is not None:
                local_keys_pressed.append(key)
                return False
            # Если это не буква, то допускаем только один модификатор
            local_keys_pressed.append(key)

    def local_on_release(key):
        return False  # Останавливаем listener

    # Создаем listener
    with keyboard.Listener(on_press=local_on_press, on_release=local_on_release) as local_listener:
        local_listener.join()  # Запускаем listener

    local_keys_pressed = keys_to_str(local_keys_pressed)
    # Обновляем текстовое поле
    field.delete(0, tk.END)
    field.insert(0, local_keys_pressed)
    return local_keys_pressed


# ==================== MONITORING МИКРОФОНА ====================
def monitor_mic():
    """
    Основной цикл мониторинга уровня звука с микрофона.
    При превышении порога эмулируется нажатие клавиш push-to-talk.
    Также реализованы функции затемнения (fade) динамиков и mute (выключение динамиков).
    """
    global mute_key_codes, ptt_keys_str, ptt_key_codes, is_talking, muted, ignore_keys_str, last_above_threshold_time, ignore_key_codes, original_speaker_volume, stored_speaker_volume, original_microphone_volume, stored_microphone_volume, pressed_keys_global

    ignore_key_codes = str_to_keys(ignore_keys_str)
    ptt_key_codes = str_to_keys(ptt_keys_str)
    mute_key_codes = str_to_keys(mute_key)

    while True:
        try:
            if mute_all_enabled and all(key in pressed_keys_global for key in mute_key_codes):
                if not muted:
                    stored_speaker_volume = volume_control.GetMasterVolumeLevelScalar()
                    volume_control.SetMasterVolumeLevelScalar(0, None)
                    stored_microphone_volume = microphone_volume.GetMasterVolumeLevelScalar()
                    microphone_volume.SetMasterVolumeLevelScalar(0, None)
                    print("Speakers and microphone muted.")
                    muted = True
                else:
                    if stored_speaker_volume is not None:
                        volume_control.SetMasterVolumeLevelScalar(stored_speaker_volume, None)
                    if stored_microphone_volume is not None:
                        microphone_volume.SetMasterVolumeLevelScalar(stored_microphone_volume, None)

                    print("Speakers and microphone unmuted.")
                    muted = False
                    stored_speaker_volume = None
                time.sleep(1)
                continue

            # Проверка активного окна
            active_window = get_active_window()
            allowed_fragments = [frag.strip().lower() for frag in allowed_window_fragments.split(',')]
            if not any(fragment in active_window for fragment in
                       allowed_fragments) and active_window != "talk to push settings":
                time.sleep(1)
                continue

            # Если нажаты клавиши для игнорирования, пропускаем обработку
            if ignore_keys_enabled and any(key in ignore_key_codes for key in pressed_keys_global):
                print("Pressing ignored")
                time.sleep(1)
                continue

            # Чтение данных с микрофона
            data = np.frombuffer(stream.read(chunk, exception_on_overflow=False), dtype=np.int16)
            current_input_level = np.mean(np.abs(data))
            update_volume_display(current_input_level)

            current_time = time.time()
            if current_input_level > volume_threshold:
                last_above_threshold_time = current_time
                if not is_talking:
                    if active_window != "talk to push settings":
                        for key in ptt_key_codes:
                            keyboard_controller.press(key)
                        if fade_sound_enabled:
                            # Сохраняем исходный уровень громкости и затемняем динамики
                            current_speaker_volume = volume_control.GetMasterVolumeLevelScalar()
                            if original_speaker_volume is None:
                                original_speaker_volume = current_speaker_volume
                            new_volume = current_speaker_volume * (1 - fade_sound_percentage / 100)
                            volume_control.SetMasterVolumeLevelScalar(new_volume, None)
                    is_talking = True
            elif is_talking and (current_time - last_above_threshold_time > post_voice_release_delay / 1000):
                if active_window != "talk to push settings":
                    for key in reversed(ptt_key_codes):
                        keyboard_controller.release(key)
                    if fade_sound_enabled and original_speaker_volume is not None:
                        volume_control.SetMasterVolumeLevelScalar(original_speaker_volume, None)
                        original_speaker_volume = None
                is_talking = False

            update_indicator()
        except OSError as e:
            print(f"Ошибка чтения с микрофона: {e}")

        time.sleep(0.01)


def update_indicator():
    """
    Обновляет цвет индикатора (лампочка) в окне настроек:
      - Зеленый, если PTT активен (речь обнаружена);
      - Красный, если PTT не активен.
    """
    if is_talking or (time.time() - last_above_threshold_time <= post_voice_release_delay / 1000):
        indicator_canvas.itemconfig(indicator_light, fill="#0DFF82")
    else:
        indicator_canvas.itemconfig(indicator_light, fill="#FF0D31")


# ==================== УСТАНОВКА МИКРОФОНА ====================
def set_microphone_device(device_index):
    """
    Останавливает предыдущий поток (если существует) и открывает новый поток для выбранного микрофона.
    """
    global stream
    if 'stream' in globals():
        stream.stop_stream()
        stream.close()
    stream = p.open(format=audio_format, channels=channels, rate=rate, input=True,
                    input_device_index=device_index, frames_per_buffer=chunk)


# ==================== ФУНКЦИИ ОКНА НАСТРОЕК И ВЫХОДА ====================
def show_settings():
    settings_window.deiconify()
    settings_window.lift()
    settings_window.focus_force()


def hide_settings():
    save_settings()
    settings_window.withdraw()


def exit_program():
    try:
        if stored_speaker_volume is not None:
            volume_control.SetMasterVolumeLevelScalar(stored_speaker_volume, None)
        if stored_microphone_volume is not None:
            microphone_volume.SetMasterVolumeLevelScalar(stored_microphone_volume, None)
    except:
        pass
    icon.stop()
    root.quit()
    root.destroy()
    if 'stream' in globals():
        stream.stop_stream()
        stream.close()
    p.terminate()
    os._exit(0)


def update_volume_display(volume):
    volume_var.set(volume)


faq_window = None  # Переменная для хранения ссылки на окно


# Функция для отображения подсказки
def show_tooltip(event, text):
    tooltip = tk.Toplevel()
    tooltip.wm_overrideredirect(True)  # Без рамок
    tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")  # Позиционируем подсказку относительно мыши

    label = ttk.Label(tooltip, text=text, relief="solid", padding=5)
    label.pack()

    # Сохраняем ссылку на тултип в объекте метки, чтобы можно было закрыть его позже
    event.widget.tooltip = tooltip


# Функция для скрытия подсказки, когда курсор покидает значок
def hide_tooltip(event):
    if hasattr(event.widget, "tooltip"):
        event.widget.tooltip.destroy()
        del event.widget.tooltip


def toggle_faq_window():
    global faq_window

    if faq_window is not None and faq_window.winfo_exists():
        faq_window.destroy()
        faq_window = None
    else:
        faq_window = tk.Toplevel()
        faq_window.title("FAQ")
        faq_window.geometry("400x150")

        # Create a frame for the canvas and scrollbar
        frame = ttk.Frame(faq_window)
        frame.pack(fill=tk.BOTH, expand=True)

        # Create a canvas widget
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # Configure the scrollable frame inside the canvas
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Attach scrollbar to canvas
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack everything
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Add FAQ content with text wrapping
        faq_text = (
            "### How It Works\n"
            "The program continuously monitors the microphone's sound level. If the level exceeds "
            "the set threshold (Limit Level), the program presses the specified Push-to-Talk keys. "
            "When you stop speaking, the keys are released after a delay (Post-Voice Delay). "
            "Additional features, such as sound fading or muting speakers, make using the program "
            "more comfortable\n\n"
            "Made by Dima_Kedr for Squad community. 2025"
        )

        label = ttk.Label(scrollable_frame, text=faq_text, wraplength=380, justify="left")
        label.pack(padx=10, pady=5)

        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

        faq_window.bind("<MouseWheel>", on_mousewheel)

        # Handle closing event
        faq_window.protocol("WM_DELETE_WINDOW", lambda: close_faq_window())


def close_faq_window():
    global faq_window
    if faq_window:
        faq_window.destroy()
        faq_window = None


# ==================== ОКНО НАСТРОЕК (TKINTER) ====================
settings_window = tk.Toplevel(root)
settings_window.title("Talk to push settings")
settings_window.geometry("400x470")
settings_window.protocol("WM_DELETE_WINDOW", hide_settings)

# Индикатор состояния
indicator_canvas = tk.Canvas(settings_window, width=20, height=20, highlightthickness=0)
indicator_canvas.pack(pady=5)
indicator_light = indicator_canvas.create_oval(2, 2, 18, 18, fill="#FF0D31")

active_volume_frame = ttk.Frame(settings_window)
active_volume_frame.pack(padx=10, pady=5, anchor="w")
# Отображение текущего уровня микрофона (только для чтения)
ttk.Label(active_volume_frame, text="Current microphone level:").pack(side="left")
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(active_volume_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Microphone's real-time volume level\n"
                                                                "and color indicator to help adjust the\nactivation "
                                                                "threshold"))
question_mark.bind("<Leave>", hide_tooltip)
current_level_scale = ttk.Scale(settings_window, from_=0, to=3000, orient='horizontal', variable=volume_var,
                                state='disabled')
current_level_scale.pack(padx=10, pady=5, fill="x")

selected_volume_frame = ttk.Frame(settings_window)
selected_volume_frame.pack(padx=10, pady=5, anchor="w")
# Ползунок для задания порога громкости
ttk.Label(selected_volume_frame, text="Limit level:").pack(side="left")
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(selected_volume_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Set the minimum volume level to trigger Push-to-Talk"))
question_mark.bind("<Leave>", hide_tooltip)
volume_threshold_scale = ttk.Scale(settings_window, from_=0, to=3000, orient='horizontal', length=250)
volume_threshold_scale.set(volume_threshold)
volume_threshold_scale.pack(padx=10, pady=5, fill="x")

# ------------------- Настройка кнопки Push-to-talk -------------------
ptt_frame = ttk.Frame(settings_window)
ptt_frame.pack(padx=10, pady=5, fill="x", anchor="w")
ttk.Label(ptt_frame, text="Push-to-talk button:        ").pack(side="left", padx=(0, 5))
ptt_key_entry = ttk.Entry(ptt_frame)
ptt_key_entry.insert(0, ptt_keys_str)
ptt_key_entry.pack(side="left", fill="x", expand=True)
ptt_edit_button = ttk.Button(ptt_frame, text="Press Key", command=lambda: get_local_pressed_keys(ptt_key_entry))
ptt_edit_button.pack(side="left", padx=(5, 0))
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(ptt_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Choose the key combination that will be pressed when"
                                                                "\nyour microphone level is above the limit (choose \n"
                                                                "the key that corresponds to the 'Push to Talk' "
                                                                "button\nin the game). Examples:"
                                                                "\n"
                                                                "\n"
                                                                "t\n"
                                                                "shift + T\n"
                                                                "f + r + y + 1 + page_down + ctrl_r"))
question_mark.bind("<Leave>", hide_tooltip)

# ------------------- Настройка игнорируемых клавиш -------------------
ignore_frame = ttk.Frame(settings_window)
ignore_frame.pack(padx=10, pady=5, fill="x", anchor="w")
ignore_keys_checkbox = tk.BooleanVar(value=ignore_keys_enabled)
ttk.Checkbutton(ignore_frame, text="Ignore if keys pressed:", variable=ignore_keys_checkbox).pack(side="left")
ignore_keys_entry = ttk.Entry(ignore_frame)
ignore_keys_entry.insert(0, ignore_keys_str)
ignore_keys_entry.pack(side="left", fill="x", expand=True)
ignore_edit_button = ttk.Button(ignore_frame, text="Press Key", command=lambda: get_local_pressed_keys(
    ignore_keys_entry))
ignore_edit_button.pack(side="left", padx=(5, 0))
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(ignore_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Prevents activating if ANY of the specified keys are\n"
                                                                "pressed, even if the microphone level exceeds\n"
                                                                "the threshold. Examples:\n\n"
                                                                "r\n"
                                                                "ctrl_l\n"
                                                                "shift + E\n"
                                                                "w + a + s + d"))
question_mark.bind("<Leave>", hide_tooltip)

# ------------------- Настройка mute -------------------
mute_frame = ttk.Frame(settings_window)
mute_frame.pack(padx=10, pady=5, fill="x", anchor="w")
mute_all_checkbox = tk.BooleanVar(value=mute_all_enabled)
ttk.Checkbutton(mute_frame, text="Mute speakers:            ", variable=mute_all_checkbox).pack(side="left")
mute_key_entry = ttk.Entry(mute_frame)
mute_key_entry.insert(0, mute_key)
mute_key_entry.pack(side="left", fill="x", expand=True)
mute_edit_button = ttk.Button(mute_frame, text="Press Key", command=lambda: get_local_pressed_keys(mute_key_entry))
mute_edit_button.pack(side="left", padx=(5, 0))
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(mute_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Set a key to quickly mute both your speakers\n"
                                                                "and microphone in case your game is interrupted\n"
                                                                "by a conversation. This prevents the game from "
                                                                "being\n"
                                                                "overheard. Works regardless of the "
                                                                "window title.\n"
                                                                "Examples:\n\n"
                                                                "+\n"
                                                                "ctrl+r\n"
                                                                "p"))
question_mark.bind("<Leave>", hide_tooltip)

# ------------------- Настройка активного окна -------------------
window_frame = ttk.Frame(settings_window)
window_frame.pack(padx=10, pady=5, fill="x", anchor="w")
ttk.Label(window_frame, text="Program name:               ").pack(side="left")
window_entry = ttk.Entry(window_frame)
window_entry.insert(0, allowed_window_fragments)
window_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(window_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Specify program window whole title\n(or its "
                                                                "fragments) where the function will work.\n"
                                                                "Examples for working in both Opera and Squad:\n"
                                                                "\n"
                                                                "opera, squad\n"
                                                                "era, squa\n"
                                                                "per, qua\n"))

question_mark.bind("<Leave>", hide_tooltip)

# ------------------- Задержка отпускания кнопки -------------------
delay_frame = ttk.Frame(settings_window)
delay_frame.pack(padx=10, pady=5, fill="x", anchor="w")
ttk.Label(delay_frame, text="Post-voice delay, ms:     ").pack(side="left")
delay_entry = ttk.Entry(delay_frame)
delay_entry.insert(0, str(post_voice_release_delay))
delay_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(delay_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Set a delay before releasing the talk keys\n"
                                                                "after you stopped speaking. This helps prevent\n"
                                                                "the last sounds from being cut off and ensures\n"
                                                                "speech isn't interrupted. Default: 800"))
question_mark.bind("<Leave>", hide_tooltip)

# ------------------- Настройка затемнения звука -------------------
fade_frame = ttk.Frame(settings_window)
fade_frame.pack(padx=10, pady=5, fill="x", anchor="w")
fade_sound_checkbox_var = tk.BooleanVar(value=fade_sound_enabled)
ttk.Checkbutton(fade_frame, text="Fade sound while talking by", variable=fade_sound_checkbox_var).pack(side="left")
fade_sound_percent_combobox = ttk.Combobox(fade_frame, values=[str(i) for i in range(101)], width=5)
fade_sound_percent_combobox.set(str(fade_sound_percentage))
fade_sound_percent_combobox.pack(side="left", padx=(5, 0))
ttk.Label(fade_frame, text="%").pack(side="left", padx=(5, 0))
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(fade_frame, text="?", cursor="hand2")
question_mark.pack(side="right", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>", lambda event: show_tooltip(event, "Reduces speaker volume while you're speaking\nto "
                                                                "avoid echo and overlapping sounds"))
question_mark.bind("<Leave>", hide_tooltip)


# ------------------- Выбор микрофона -------------------
frame = ttk.Frame(settings_window)
frame.pack(padx=10, pady=5, fill="x", anchor="w")
ttk.Label(frame, text="Choose active microphone:").pack(side="left")
mic_choice_var = tk.StringVar()
# Добавление значка вопроса с тултипом
question_mark = ttk.Label(frame, text="?", cursor="hand2")
question_mark.pack(side="left", padx=(5, 0))
# Добавление событий для появления и скрытия всплывающего окна
question_mark.bind("<Enter>",
                   lambda event: show_tooltip(event, "Select the microphone if you have multiple microphones."))
question_mark.bind("<Leave>", hide_tooltip)
mic_choice_menu = ttk.Combobox(settings_window, textvariable=mic_choice_var, state='readonly')
mic_choice_menu['values'] = [mic[1] for mic in microphones]
if 0 <= selected_mic_index < len(microphones):
    mic_choice_menu.set(microphones[selected_mic_index][1])
else:
    mic_choice_menu.set(microphones[0][1])
mic_choice_menu.pack(padx=10, pady=5, fill="x")


# ------------------- Кнопки управления -------------------
bottom_button_frame = ttk.Frame(settings_window)
bottom_button_frame.pack(padx=10, pady=10, fill="x")
ttk.Button(bottom_button_frame, text="FAQ", command=toggle_faq_window).pack(side="left", padx=5)
ttk.Button(bottom_button_frame, text="OK", command=hide_settings).pack(side="right", padx=5)
ttk.Button(bottom_button_frame, text="Apply", command=save_settings).pack(side="right", padx=5)

settings_window.withdraw()  # Скрываем окно настроек по умолчанию

# ==================== ТРЕЙ-ИКОНКА ====================
# Base64 строка изображения (вставьте сюда строку, полученную на предыдущем шаге)
encoded_icon = "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"

# Декодируем base64 строку в бинарные данные
icon_data = base64.b64decode(encoded_icon)

# Используем BytesIO для преобразования в файл
icon_image = Image.open(BytesIO(icon_data))

# Теперь можно использовать картинку как обычно
tray_menu = (item('Settings', show_settings), item('Exit', exit_program))
icon = Icon("MicTrigger", icon_image, menu=tray_menu)

# ==================== ЗАПУСК ПРОГРАММЫ ====================
set_microphone_device(microphones[selected_mic_index][0])
monitor_thread = threading.Thread(target=monitor_mic, daemon=True)
monitor_thread.start()


def start_pystray():
    icon.run()


pystray_thread = threading.Thread(target=start_pystray, daemon=True)
pystray_thread.start()

root.mainloop()
