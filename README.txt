COLEAT = Collabora OLE Automation Translator

    IMPORTANT: Note that for the redirection to Collabora Office to
    work, you need a version of Collabora Office with changes that
    aren't available yet in any build of it. Always use the -n option.

COLEAT (coleat.exe) is an application that runs another application
(passed on its command line) in one of two modes: Either as such, but
with various amounts of logging and tracing output, or redirecting the
app's use of Word and Excel COM components to Collabora Office ones
instead.

Terminology:
    The program being run under COLEAT is called the "wrapped client
    application".

    The application offering Automation or COM services that the
    client application is supposed to use is called the "original
    application".

    The one that the client application gets redirected to instead is
    the "replacement application".

(Even though I use Word, Excel, and Collabora Office here as examples,
COLEAT could also be built to work against some other OLE Automation
and COM service offering application, and redirect use of that to
another replacement application.)

With no options, COLEAT does just the redirection functionality. I.e.
if you run it on a client application (written in VB6, C++, VBScript,
or some other programming language) that starts Word to open a
document, and do something on it, COLEAT will make it instead start
Collabora Office and open the document in that, and do the
corresponding operations on the document. Note: For now, the
functionality available in Collabora Office that matches that in Word
closely enough is fairly limited.

With the -n option, no redirection takes place, and the client
application should work as it does without being wrapped by COLEAT.
The only difference is that by also using the -t option, you will get
tracing output describing the API of the target app (Word or Excel)
used.

There is also an option -v that gives verbose output than -t, but it
is mostly intended as a debugging tool for COLEAT itself.

COLEAT needs to be installed so that the three .exe files and two .dll
files are in the same folder.

Summary: Use it like this, in a Command Prompt window:

coleat -n -t cscript demo.vbs

to run the included demo.vbs VBScript program, that opens a couple of
documents in Word and does some (very) simple things with them.

or:

coleat -n -t demo.exe

to run the included demo.exe VB6 program, that does the same.
