import sys, os, json
from monitorcontrol import PowerMode, get_monitors
from pythonosc import osc_server
from pythonosc.dispatcher import Dispatcher
from threading import Thread, Event
from time import sleep, strftime
from traceback import format_exc
from infi.systray import SysTrayIcon
import tkinter as tk
from tkinter import ttk, filedialog

BACKUP_LOGFILE = "monitor_osc.log"

# TODO: monitorcontrol has an issue where you can't uniquely identify two monitors with the same model code.
#  https://github.com/newAM/monitorcontrol/issues/250

# TODO: need to detect when a monitor is turned off and back on.

class Monitor:
    def __init__(self, manager, monitor):
        """
        Controls the state of a single monitor.
        Basically a wrapper around the monitorcontrol Monitor object.
        :param manager: the manager object
        :param monitor: the monitorcontrol Monitor object
        """
        self.manager = manager  # reference to the manager
        self.monitor = monitor  # core monitor object
        self.model = None       # model name of the monitor
        self.running = False

        # get core monitor object from monitorcontrol
        try:
            with self.monitor:
                data = self.monitor.get_vcp_capabilities()
                model = data.get("model")
                if model:
                    self.model = model
                else:
                    self.debug(f"Could not identify model of a monitor: {data}")
        except Exception as e:
            self.debug(f"Problem getting monitor info.\n{e.__class__.__name__}: {e}\n\n{format_exc()}")

        # Queued values
        self.contrast = None
        self.luminance = None

        # Property ranges
        self.contrast_range = (0, 100)
        self.luminance_range = (0, 100)

        # Offset ranges
        self.contrast_offset = (0, 1)
        self.luminance_offset = (0, 1)

        # event to signal when to start setting queued values
        self.event = Event()

        # Interval between monitor updates (ms)
        self.interval = 10

        # # detect when monitor is turned off
        # self.last_power_mode = self.monitor.get_power_mode()
        # self.last_power_time = time()

    def debug(self, *msg):
        self.manager.debug(*msg)

    def run(self):
        """ Run the event loop for this monitor, waiting for signals to set contrast/luminance. """
        if self.running: return  # do nothing if already running
        self.running = True
        self.manager.log("Running monitor:", self.model)
        Thread(target=self._run).start()

    def _run(self):
        """
        Must be run on separate thread.
        Run the event loop for this monitor, waiting for signals to set contrast/luminance.
        This sets the contrast/luminance of the monitor according to the most recently queued values,
            and waits according to the interval before checking again.
        """
        while self.running:
            self.event.wait()  # wait to be notified of a signal
            self.verify_monitor()

            while self.event.is_set():  # while commands need to be executed
                self.event.clear()  # reset event. This will only loop if more values are immediately queued

                try:
                    self.set_luminance()
                    self.set_contrast()
                except Exception as e:
                    self.manager.log(f"Error setting monitor values: {self.model}.\n{e.__class__.__name__}: {e}\n\n{format_exc()}")

    def stop(self):
        """ Stop the monitor event loop """
        self.running = False

    def verify_monitor(self):
        """ Run periodically to verify the monitor is still connected """
        try:
            with self.monitor:
                pass
        except Exception as e:
            self.manager.log(f"Error connecting to monitor: {self.model}\n{e.__class__.__name__}: {e}")
            self.manager.locate_monitors()  # re-locate monitors

    def set_interval(self, interval):
        self.interval = interval or self.interval

    # Bind commands
    def bind_contrast(self, path, range=None, offset=None):
        """
        Bind the monitor contrast to this path.
        :param path: string OSC path to bind
        :param range: tuple for output range. Default (0, 100).
            What the min and max output values are.
        :param offset: tuple for output range offset. Default (0, 1).
            Where hte min and max output values are relative to the input values.
        """
        if range: self.contrast_range = range
        if offset: self.contrast_offset = offset
        self.manager.bind(path, self.queue_contrast)

    def bind_luminance(self, path, range=None, offset=None):
        """
        Bind the monitor luminance to this path.
        :param path: string OSC path to bind
        :param range: tuple for output range. Default (0, 100).
            What the min and max output values are.
        :param offset: tuple for output range offset. Default (0, 1).
            Where hte min and max output values are relative to the input values.
        """
        if range: self.luminance_range = range
        if offset: self.luminance_offset = offset
        self.manager.bind(path, self.queue_luminance)

    def bind_toggle(self, path):
        """ Bind the monitor toggle to this path """
        self.manager.bind(path, self.toggle)

    # Queue signals
    def queue_luminance(self, value):
        """ Queue a luminance value """
        self.luminance = value
        self.event.set()  # Signal an update

    def queue_contrast(self, value):
        """ Queue a contrast value """
        self.contrast = value
        self.event.set()  # Signal an update

    # Set monitor values
    def set_luminance(self):
        """ Set the monitor luminance """
        if self.luminance is None: return
        value = (self.luminance - self.luminance_offset[0]) / (self.luminance_offset[1] - self.luminance_offset[0])
        value = (value * (self.luminance_range[1] - self.luminance_range[0])) + self.luminance_range[0]
        value = min(max(value, self.luminance_range[0]), self.luminance_range[1])
        with self.monitor:
            self.monitor.set_luminance(int(value))
        sleep(self.interval / 1000)

    def set_contrast(self):
        """ Set the monitor contrast """
        if self.contrast is None: return
        value = (self.contrast - self.contrast_offset[0]) / (self.contrast_offset[1] - self.contrast_offset[0])
        value = (value * (self.contrast_range[1] - self.contrast_range[0])) + self.contrast_range[0]
        value = min(max(value, self.contrast_range[0]), self.contrast_range[1])
        with self.monitor:
            self.monitor.set_contrast(int(value))
        sleep(self.interval / 1000)

    def toggle(self, value):
        """ Toggle the monitor on and off """
        if value == 0: return  # ignore up-press
        with self.monitor:
            state = self.monitor.get_power_mode()
            if state == PowerMode.on:
                self.monitor.set_power_mode(PowerMode.off_soft)
            else:
                self.monitor.set_power_mode(PowerMode.on)


class Manager:
    def __init__(self, debug=False):
        """
        :param ip: Local IP address for OSC
        :param port: Port for OSC. If not specified, must specify for each mapping.
        """
        self.root = None  # tk root
        self.running = True
        self.sections = []  # reset on open(), holds the tk section frames

        self.ip = "127.0.0.1"
        self.port = 5000
        self.monitor_settings = []
        self.debug_mode = debug
        self.server = None  # the OSC server
        self.tray = None  # the system tray

        # Set path to assets directory
        self.asset_dir = ""
        self.get_asset_path()
        self.icon = f"{self.asset_dir}/monitor.ico"

        self.log_file = os.path.join(self.asset_dir, BACKUP_LOGFILE)
        self.save_file = os.path.join(self.asset_dir, "config.json")

        # Load settings from disk
        self.load()

        # List of Monitor() objects
        self.monitors = []
        self.locate_monitors()

        # Map of paths to lists of callback functions
        self.paths = {}

    def log(self, *args):
        msg = f"[{strftime('%Y-%m-%d %H:%M:%S')}] {' '.join(args)}"
        log_file = BACKUP_LOGFILE
        if hasattr(self, "log_file") and os.path.exists(self.log_file):
            log_file = self.log_file

        print(msg)
        with open(log_file, "a") as f:
            f.write('\n'+msg)
    def debug(self, *args):
        if not self.debug_mode: return
        self.log("[DEBUG]", *args)
    def get_asset_path(self):
        """ Get path to assets directory """
        # first check if script is running as a python file or an executable
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(__file__)

        self.asset_dir = os.path.join(application_path, "assets")

    ######
    # Monitor control
    def locate_monitors(self, *args):
        """ Get all connected monitors and store them as Monitor() objects """
        self.log("Identifying Monitors...")
        monitors = get_monitors()  # get monitor objects from monitorcontrol
        self.log(f"Found {len(monitors)} monitors...")
        for monitor in monitors:
            m = Monitor(self, monitor)

            # first check if monitor is already in list
            for existing in self.monitors:
                if existing.model == m.model:
                    self.log(f"Identified existing monitor: {m.model}")
                    existing.monitor = monitor  # update monitor object
                    break
            else:  # otherwise add it to the list
                self.log(f"Identified new monitor: {m.model}")
                self.monitors.append(m)
        return self.monitors
    def get_monitor(self, model):
        """ Get a monitor my model """
        for monitor in self.monitors:
            if monitor.model == model:
                monitor.run()  # start the event loop for this monitor (if it isn't already)
                return monitor
        self.log(f"Couldn't find monitor with model: {model}")
    def bind(self, path, func):
        """ Add a callback function to a path """
        if not self.paths.get(path):
            self.paths[path] = []
        self.paths[path].append(func)
        self.debug(f"Bound: \"{path}\"")
    def trigger(self, unused, args, value):
        """
        This is the main callback for OSC signals.
        :param args: first arg is a list of functions to trigger.
        :param value: value of the OSC signal to be passed to each function
        """
        functions = args[0]
        for function in functions:
            try:
                function(value)
            except Exception as e:
                self.log(f"Error in function: {function.__name__}\n{e.__class__.__name__}: {e}\n\n{format_exc()}")
    def set_configuration(self):
        """ Set the monitor configuration based on the interface settings """
        self.paths = {}  # reset paths

        self.log("Setting config from interface")
        for settings in self.monitor_settings:
            monitor = self.get_monitor(settings.get('id'))
            if not monitor: continue

            interval = settings.get('interval')
            monitor.set_interval(interval)

            contrast_path = settings.get('contrast_path')
            lum_path = settings.get('lum_path')
            toggle_path = settings.get('toggle_path')

            contrast_range = (settings.get('contrast_range_min'), settings.get('contrast_range_max'))
            contrast_offset = (settings.get('contrast_offset_min'), settings.get('contrast_offset_max'))

            lum_range = (settings.get('lum_range_min'), settings.get('lum_range_max'))
            lum_offset = (settings.get('lum_offset_min'), settings.get('lum_offset_max'))

            if contrast_path:
                monitor.bind_contrast(contrast_path, range=contrast_range, offset=contrast_offset)
            if lum_path:
                monitor.bind_luminance(lum_path, range=lum_range, offset=lum_offset)
            if toggle_path:
                monitor.bind_toggle(toggle_path)

    def run(self):
        """ Run continuously """
        # start the tray icon
        self.debug("Setting up tray icon")
        self.tray = SysTrayIcon(self.icon, "DDCCI Monitor Manager", (("Config", None, self.open),), on_quit=self.quit)
        if self.tray:
            Thread(target=self.tray.start).start()

        # Continually run, reloading when killed
        while self.running:
            self.set_configuration()  # set the OSC bindings from the current configuration

            # set the binding paths
            dispatcher = Dispatcher()
            self.debug("PATHS: ", str(self.paths))
            for path, functions in self.paths.items():
                dispatcher.map(path, self.trigger, functions)

            self.server = osc_server.ThreadingOSCUDPServer((self.ip, self.port), dispatcher)
            self.log("Serving on {}".format(self.server.server_address))
            self.server.serve_forever()
            self.debug("Server Stopped")
    def reload(self):
        """ Reload the OSC server """
        if not self.server: return
        self.debug("Reloading OSC server...")
        self.server.shutdown()  # stop the server if it's already running
        self.server.server_close()
        self.server = None
    def quit(self, *args):
        """ Quit all threads """
        self.running = False
        sleep(1)
        try:
            exit()
        except:
            os._exit(0)  # force

    ##############
    # Interface
    def save(self):
        self.ip = self.ip_var.get()
        self.port = self.port_var.get()
        self.log_file = self.log_file_var.get()
        self.monitor_settings = []

        for section in self.sections:
            self.monitor_settings.append({
                "id": section['id_var'].get(),
                "interval": section['interval'].get(),
                "toggle_path": section['toggle_path'].get(),
                "contrast_path": section['contrast_path'].get(),
                "contrast_range_min": section['contrast_range'][0].get(),
                "contrast_range_max": section['contrast_range'][1].get(),
                "contrast_offset_min": section['contrast_offset'][0].get(),
                "contrast_offset_max": section['contrast_offset'][1].get(),
                "lum_path": section['lum_path'].get(),
                "lum_range_min": section['lum_range'][0].get(),
                "lum_range_max": section['lum_range'][1].get(),
                "lum_offset_min": section['lum_offset'][0].get(),
                "lum_offset_max": section['lum_offset'][1].get(),
            })
        data = {
            "ip": self.ip,
            "port": self.port,
            "monitors": self.monitor_settings,
            "log_file": self.log_file,
            "save_file": self.save_file,
        }

        try:
            with open(self.save_file, "w") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.log("Error saving config:", e)
    def load(self):
        if not os.path.exists(self.save_file):
            return

        try:
            with open(self.save_file, "r") as f:
                data = json.load(f)
                self.ip = data.get('ip', self.ip)
                self.port = data.get("port", self.port)
                self.monitor_settings = data.get("monitors", self.monitor_settings)
                self.log_file = data.get("log_file", self.log_file)
        except Exception as e:
            self.log("Error loading config:", e)
    def open(self, tray):
        if self.root:
            self.root.lift()
            return

        # create tk interface
        self.build_interface()
        self.populate_dropdowns()
        self.root.mainloop()

    def build_interface(self):
        self.root = tk.Tk()
        self.root.title("Configuration")
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.geometry("600x400")

        # Updates to the tk interface must be on the main thread, so we register a main thread event that background threads can trigger.
        self.root.bind("<<populate_dropdowns>>", self._populate_dropdowns)

        self.ip_var = tk.StringVar(master=self.root, value=self.ip)
        self.port_var = tk.IntVar(master=self.root, value=self.port)
        self.log_file_var = tk.StringVar(master=self.root, value=self.log_file)
        self.sections = []

        # Log File
        tk.Label(self.root, text="Log File").grid(row=0, column=0, sticky="e")
        tk.Entry(self.root, textvariable=self.log_file_var, width=40).grid(row=0, column=1, sticky="ew")
        tk.Button(self.root, text="Browse...", command=self.browse_log_file).grid(row=0, column=2, padx=5)

        # IP, Port and buttons
        tk.Label(self.root, text="IP").grid(row=1, column=0, sticky="e")
        tk.Entry(self.root, textvariable=self.ip_var).grid(row=1, column=1, sticky="w")

        tk.Label(self.root, text="Port").grid(row=2, column=0, sticky="e")
        tk.Entry(self.root, textvariable=self.port_var).grid(row=2, column=1, sticky="w")

        def _save():
            self.save()
            self.set_configuration()
            self.reload()

        def _refresh():
            self.populate_dropdowns(force_reload=True)

        f = tk.Frame(self.root)
        f.grid(row=3, column=0, columnspan=5, sticky="w")
        tk.Button(f, text="Save", command=_save).pack(side="left", padx=5, pady=5)
        tk.Button(f, text="Refresh", command=_refresh).pack(side="left", padx=5, pady=5)
        tk.Button(f, text="Add Monitor", command=self.add_section).pack(side="left", padx=5, pady=5)

        # Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=4, column=0, columnspan=5, pady=10, sticky="ew")

        # Set col 1 to expand
        self.root.columnconfigure(1, weight=1)

        if not self.monitor_settings:
            self.add_section()
        else:
            for data in self.monitor_settings:
                self.add_section(prefill=data)
    def add_section(self, prefill=None):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=f"Monitor {len(self.sections) + 1}")

        section = {}
        row = 0  # int to keep track of the current row

        # ID Dropdown
        id_var = tk.StringVar(master=self.root, value=prefill.get("id") if prefill else "")
        tk.Label(frame, text="ID").grid(row=row, column=0, sticky="w")
        id_menu = ttk.OptionMenu(frame, id_var, id_var.get(), *[m.model for m in self.monitors])
        id_menu.grid(row=row, column=1, sticky="w")
        section['id_var'] = id_var
        section['id_menu'] = id_menu

        # Delete button
        def delete_this_section():
            self.delete_section(section, frame)
        tk.Button(frame, text="Delete", fg="red", command=delete_this_section).grid(row=row, column=2, sticky="e", padx=10)

        # Interval (integer ms)
        row += 1
        interval_var = tk.IntVar(master=self.root, value=prefill.get("interval", 10) if prefill else 10)
        tk.Label(frame, text="Interval (ms)").grid(row=row, column=0, sticky="w")
        tk.Entry(frame, textvariable=interval_var, width=5).grid(row=row, column=1, sticky="w")
        section['interval'] = interval_var

        # Toggle Path
        row += 1
        tk.Label(frame, text="Toggle Path").grid(row=row, column=0, sticky="w")
        toggle_path = tk.Entry(frame)
        toggle_path.grid(row=row, column=1, sticky="ew")
        if prefill: toggle_path.insert(0, prefill.get("toggle_path", ""))
        section['toggle_path'] = toggle_path

        # Contrast
        row += 1
        tk.Label(frame, text="Contrast Path").grid(row=row, column=0, sticky="w")
        contrast_path = tk.Entry(frame)
        contrast_path.grid(row=row, column=1, sticky="ew")
        if prefill: contrast_path.insert(0, prefill.get("contrast_path", ""))
        section['contrast_path'] = contrast_path

        contrast_range = [
            tk.IntVar(master=self.root, value=prefill.get("contrast_range_min", 0) if prefill else 0),
            tk.IntVar(master=self.root, value=prefill.get("contrast_range_max", 100) if prefill else 100)
        ]
        contrast_offset = [
            tk.DoubleVar(master=self.root, value=prefill.get("contrast_offset_min", 0.0) if prefill else 0.0),
            tk.DoubleVar(master=self.root, value=prefill.get("contrast_offset_max", 1.0) if prefill else 1.0)
        ]

        row += 1
        tk.Label(frame, text="Contrast Range (min/max)").grid(row=row, column=0, sticky="w")
        f = tk.Frame(frame)
        f.grid(row=row, column=1, sticky="w")
        tk.Entry(f, textvariable=contrast_range[0], width=5).pack(side="left")
        tk.Entry(f, textvariable=contrast_range[1], width=5).pack(side="right")

        row += 1
        tk.Label(frame, text="Contrast Offset (min/max)").grid(row=row, column=0, sticky="w")
        f = tk.Frame(frame)
        f.grid(row=row, column=1, sticky="w")
        tk.Entry(f, textvariable=contrast_offset[0], width=5).pack(side="left")
        tk.Entry(f, textvariable=contrast_offset[1], width=5).pack(side="right")

        section['contrast_range'] = contrast_range
        section['contrast_offset'] = contrast_offset

        # Luminescence
        row += 1
        tk.Label(frame, text="Luminescence Path").grid(row=row, column=0, sticky="w")
        lum_path = tk.Entry(frame)
        lum_path.grid(row=row, column=1, sticky="ew")
        if prefill: lum_path.insert(0, prefill.get("lum_path", ""))
        section['lum_path'] = lum_path

        lum_range = [
            tk.IntVar(master=self.root, value=prefill.get("lum_range_min", 0) if prefill else 0),
            tk.IntVar(master=self.root, value=prefill.get("lum_range_max", 100) if prefill else 100)
        ]
        lum_offset = [
            tk.DoubleVar(master=self.root, value=prefill.get("lum_offset_min", 0.0) if prefill else 0.0),
            tk.DoubleVar(master=self.root, value=prefill.get("lum_offset_max", 1.0) if prefill else 1.0)
        ]

        row += 1
        tk.Label(frame, text="Luminescence Range (min/max)").grid(row=row, column=0, sticky="w")
        f = tk.Frame(frame)
        f.grid(row=row, column=1, sticky="w")
        tk.Entry(f, textvariable=lum_range[0], width=5).pack(side="left")
        tk.Entry(f, textvariable=lum_range[1], width=5).pack(side="right")

        row += 1
        tk.Label(frame, text="Luminescence Offset (min/max)").grid(row=row, column=0, sticky="w")
        f = tk.Frame(frame)
        f.grid(row=row, column=1, sticky="w")
        tk.Entry(f, textvariable=lum_offset[0], width=5).pack(side="left")
        tk.Entry(f, textvariable=lum_offset[1], width=5).pack(side="right")

        section['lum_range'] = lum_range
        section['lum_offset'] = lum_offset

        # Set col 1 to expand
        frame.columnconfigure(1, weight=1)
        self.sections.append(section)
    def delete_section(self, section, frame):
        frame.destroy()
        self.sections.remove(section)
    def close(self):
        self.save()
        self.root.quit()
        self.root.destroy()
        self.root = None
        self.reload()

    def populate_dropdowns(self, force_reload=False):
        """
        Populate the monitor selection dropdowns from a background thread.
        If no monitors currently located or force_reload=True, relocates all monitors.
        """
        def thread():
            if force_reload or len(self.monitors) == 0:
                self.locate_monitors()
            self.root.event_generate("<<populate_dropdowns>>")
        Thread(target=thread).start()
    def _populate_dropdowns(self, event=None):
        """ Populates the monitor selection dropdowns. Must be run from the main thread. """
        for section in self.sections:
            menu = section['id_menu']
            menu["menu"].delete(0, "end")
            for monitor in self.monitors:
                menu["menu"].add_command(label=monitor.model, command=tk._setit(section['id_var'], monitor.model))

    def browse_log_file(self):
        path = filedialog.askopenfilename(title="Select Log File")
        if path:
            self.log_file_var.set(path)


if __name__ == "__main__":
    try:
        manager = Manager()
        manager.run()
    except Exception as e:
        msg = f"\nError occurred.\n\n{e.__class__.__name__}: {e}\n\n{format_exc()}"
        with open(BACKUP_LOGFILE, "a") as f:
            f.write(msg)
        print(msg)