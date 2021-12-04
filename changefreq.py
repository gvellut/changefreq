import pywintypes  # noqa
import win32api
import win32con


def change_freq(m, freq):
    mi = win32api.GetMonitorInfo(m[0])
    devmode = win32api.EnumDisplaySettings(
        mi["Device"], win32con.ENUM_REGISTRY_SETTINGS
    )
    devmode = pywintypes.DEVMODEType()
    devmode.DisplayFrequency = freq
    devmode.Fields = win32con.DM_DISPLAYFREQUENCY
    retcode = win32api.ChangeDisplaySettingsEx(
        mi["Device"],
        devmode,
        win32con.CDS_UPDATEREGISTRY | win32con.CDS_RESET,
    )
    return retcode


def list_monitors():
    monitors = win32api.EnumDisplayMonitors()
    return monitors


def list_freqs(m):

    freqs = set()
    mi = win32api.GetMonitorInfo(m[0])

    # to get the current res
    devmode = win32api.EnumDisplaySettings(
        mi["Device"], win32con.ENUM_REGISTRY_SETTINGS
    )
    res_w = devmode.PelsWidth
    res_h = devmode.PelsHeight

    # loop equivalent to "List All Modes" in "Display Adapter Properties for Display"
    imode = 0
    while True:
        try:
            devmode = win32api.EnumDisplaySettings(mi["Device"], imode)
        except Exception:
            break

        # only list freqs valid at current res
        if devmode.PelsWidth == res_w and devmode.PelsHeight == res_h:
            freqs.add(devmode.DisplayFrequency)
        imode += 1
    return freqs


if __name__ == "__main__":
    ms = list_monitors()
    print(f"{len(ms)} screens")

    for m in ms:
        freqs = list_freqs(m)
        max_freq = max(freqs)
        dummy_freq = next(filter(lambda x: x != max_freq, freqs))
        # just in case ; CDS_RESET flag should be enough
        change_freq(m, dummy_freq)
        retcode = change_freq(m, max_freq)

        if retcode == win32con.DISP_CHANGE_SUCCESSFUL:
            print(f"!!! OK : {max_freq} Hz")
        else:
            print(f"!!! Not OK : {max_freq} Hz")
