# pio
Peek in Outlook. 

Powershell script that peeks inside Outlook inbox and displays last messages. 

### Features:
- Marks messages that contain specific string in `To` field by adding `>` in the leftmost column.
- Highlights unread messages (White color)
- Runs in infinite while loop
- Sleeps for x seconds after every iteration
- adapts to terminal size based on host `ui.rawui.windowsize` (only at the begining of iteration)
- shows inbox parsing progress by displaying poor-man's progress bar `-\|/-`


### Censored (blured) example

![Peek in Outlook](https://raw.githubusercontent.com/chaoos-theory/pio/master/pio_example.jpg)
