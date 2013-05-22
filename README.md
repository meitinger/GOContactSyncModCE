GO Contact Sync Mod CE
======================
This is a fork of (*GO Contact Sync Mod*)[http://sourceforge.net/projects/googlesyncmod/]
which in turn is a fork of the no longer maintained *GO Contact Sync* project.
The *CE* stands for *corporate edition*, which shouldn't imply any commercial
intentions, but rather what kind of end users this fork is tailored to.

What's different?
-----------------
Basically, the UI is rewritten from scratch to make it look like *Google
Calendar Sync*, i.e. something that most users are familiar with and doesn't
allow too much tinkering by the user.
It is also more standards-based, since it facilitates data binding, ordinary
.NET application settings and background workers. The notification icon is now
the same as the app icon and is animated during synchronization. Balloon hints
are only shown if the sync was started interactively (although can be reviewed
by clicking on the notification icon), error messages are only displayed if the
operation fails unexpectedly. Everything else is written to the log file which
can be shown through the notification icon's menu.
Apart from the UI, only minor corrections/alternations have been made, perhaps
most significantly to logging, since the original project dumped the file into
the roaming app data instead of the local app data.
The installer has also been replaces by a WiX project.

What about support?
-------------------
If you need support for synchronization, I strongly recommend the original
project. Although the changes to the core classes are minimal, there is a good
chance that no one from the original project will either be able or willing to
help you. Then again, it might be easier to support this project yourself if
you are responsible for deployment and know how to program.
The long-term plan is to keep up with the SVN master and merge any relevant
changes into the core classes within the corporate branch.
