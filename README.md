# Neat Button Scan

Small script (with a nice GUI) for scanning triggered by scanner's button press. For Windows (10) only.
It works well on my PC, but bugs are expected due to different environments. Please report if you find one.

## What's required?

- a scanner with a button(s)
- [PowerShell](https://github.com/PowerShell/PowerShell) & [AutoHotkey](https://www.autohotkey.com)

## How it works?

There are two scripts: one `nbs.ahk` ([AutoHotkey](https://www.autohotkey.com)), and one `nbs.ps1` ([PowerShell](https://github.com/PowerShell/PowerShell)). The first one show a little GUI, detects button press, and calls the second one which does the scanning. So fire up `nbs.ahk` and press a scanner's button. To close the "app" press `Shift+Esc`.

### Button press detection

If there is no default scanner's button action set, Windows shows a popup window for you to select an action to perform (execute an app). Neat Button Scan detects this little window, immediately closes it, and then starts its job. So if you *have* set a default app for the button you press, Neat Button Scan will *not* work – the default app will be executed instead. Go to Windows' scanner setting to set/reset default action for a button you want to use.

## Configure and test

There is no separate configuration file. Both scripts have configurable variables on the top. So you have to modify them in order to use it.
To check scanning is working, execute the [PowerShell](https://github.com/PowerShell/PowerShell)'s script by hand:

```
$ .\nbs.ps1
```

Scanner should start scanning and you should see, in the end, something like:

```
Scanning to D:\Home\Documents\From-Scanner\img-023.tiff ...
Done!!
```

If not, and you're sure your settings are correct, please report a bug.

## Screenshots

![Ready](https://raw.githubusercontent.com/spajak/neat-button-scan/master/screenshots/ready.png)

![Scanning](https://raw.githubusercontent.com/spajak/neat-button-scan/master/screenshots/scanning.png)

![Waiting](https://raw.githubusercontent.com/spajak/neat-button-scan/master/screenshots/waiting.png)

## TODO

- Separate configuration file
- Error handling
