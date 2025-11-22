# pipe2textbox
[download binary](https://github.com/samtupy/pipe2textbox/releases/latest/download/show.exe)

Now it's command|show, no more command|clip / run > notepad / paste!

## What is this thing?
If you are an NVDA screen reader user who also employs the command line, no doubt you will have run into the infuriating issue where only the very tail portion of a command's output is easily accessible to you. How many times have you had to type "command --help > help.txt" or "command --help | clip" before pasting into notepad, all because NVDA can't scroll up in the CLI output history? You can use control+A then control+c in windows cmd, but you still must sacrifice or back up your clipboard contents for this not to mention opening notepad to review the data. Seriously if you wanted to read more than the last 30ish lines of terminal output, until now it always somehow had to involve opening notepad!

While this little program I've made by no means improves NVDA's scrolling functionality itself, it does attempt to make the entire situation just a little less painful by allowing you to pipe the output of any command into a simple multi-line read only text box for easy review. This means that as an NVDA user, you can now review the long output of a process without needing to sacrifice your clipboard's contents, without having to create some temporary file on disk, without having to launch notepad or generally just without any extra 5 seconds of hastle where sighted people spend 0! It even includes a find dialog (ctrl+f, f3, shift+f3) which means that you don't even need to copy the text to notepad to search for a specific phrase! And if you decide you want to preserve a command's output, just press ctrl+s on the text field and you'll be covered!

## Usage
It really couldn't be more simple, just drop show.exe somewhere on your path (anywhere from c:\windows to c:\users\%username% so long as it's part of the path environment variable), then in the command line, you can test by running "dir c:\windows\system32|show" and see how you can instantly review every bit of that massive directory listing where as before you could only view the last 30 lines of it without extra hastle! Of course you can pipe anything to show.exe that you could pipe to the clip command. Easily reviewing git logs in a seekable manner is another great use for this, for example.

The resulting dialog is very simple, containing a read only edit box with command output, a find button, and a save as button. You can also press ctrl+f or f3 in the text field to open the standard find dialog. As with most applications, f3 and shift+f3 remember the last search term so that you can jump between occurances of text easily. The finding here doesn't wrap around yet, I'll likely get around to that soon. The save as feature can also be accessed by pressing control+s in the text field. The aplication always saves files encoded in UTF8.

You can exit the application by pressing either escape or alt+f4. You can also press ctrl+c in the terminal window you launched the application from by piping another process to it.

It's not any more complicated than that!

## Limitations
The main thing to warn about here is that for the first iteration of this program, the same behavior of the clip command is used (E. the text from a piping process queues up, then the text box displays the text once the piping process exits). This means that this isn't yet suitable for reviewing live output from long running CLI programs (say ssh or python). In the future, I'll make a new version of this thing that reads the piped text on a separate thread and continually updates the text field as the piping process outputs more text. This however requires several other considerations (such as screen reader speech to announce live output), so I've forgone it for the moment in light of at least releasing an initial version of this thing as it's already quite useful as is.

I've done my very best under the circumstances to support UTF16, UTF8, system codepage and windows-1252 files for reading if they are piped in to the program. I do not know if this will be sufficient in all cases, please let me know if you find a plain text file that easily opens in your text editor but can't be successfully piped to the program.

## Building
CMake is used as the build system for this program, targeting MSVC 2022. It's possible that older versions of Visual Studio will work as well, but that seemingly broke the ability to make raw COM calls from pure C from some basic testing.

After cloning the repository, simply running:
```
mkdir build && cd build
cmake .. && cmake --build .
```

will give you a fully functional and tiny show.exe in the build directory.

## So much code for such a simple program, why?
Trust me, I'm well aware that I have written this program at a rediculisly low level, it could probably be spun up in 40 lines of python and not much more c# or something. However, I was already interested in seeing how small of an executable Visual Studio could generate without a c runtime library, and my idea was to create a simple program. Indeed, the initial version of this thing was less than 90 lines of code, contained 3 functions, and could have possibly been made even smaller. But then I decided to add a find dialog which added another 50 odd lines, then decided to throw in a save feature as well so there's another 50ish, then I decided right before publishing the project that I wanted to support more text file encodings if someone pipes a file to the program, that was another chunk of more than 50 lines of code. In the end, I ended up learning a lot more than I set out to during this development process, however have happily reached my side goal utterly as compiling with vs2008 produces an executable that's only 6.5kb! The ssl signing on the official release I provide shoots it up to 12k instead. All things considered, in the end I'm not unhappy. It took literally days to develop this, there were some painful moments where I was clearly reminded how much easier this would have been had I done this at a higher level, but certainly the fun and challenge made it well worth it for me. If non-c-programmers have ideas for this thing and want to contribute, I may indeed rewrite it at a higher level. For now though, I hope you all get as much use out of this little program as I will get from the knowledge I've collected while making it!
