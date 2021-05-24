
This is a short
[AutoHotkey](https://www.autohotkey.com)
script to take raw events from a
[VEC Infinity USB-3 footpedal](http://veccorp.com/foot-controls.html)
and translate them into page-up / page-down events to powerpoint.

(My use-case is to be able to advance powerpoint slides while running a webinar without having powerpoint focussed)

In its current form: its unlikely to be useful to anyone else in its raw form.

I'm publishing it as an example of how to collect events from the footpedal and dispatch them to other applications.

This script depends on the
[AHKHID module](https://github.com/jleb/AHKHID)
to do all the heavy lifting.  It is packaged as a submodule.
If you have not yet cloned this module you can clone it by:

```shell
git clone --recurse-submodules https://github.com/neilkatin/pptvec
```

If you've already cloned this repo you need to initialize the AHKHID submodule (from the pptvec folder):

```shell
git submodule init
git submodule update
```

The script has a very rudimentary debug system in it to display debug messages.
Calling the script with the ```--debug``` argument will output messages to stderr.
Calling the script with the ```---debugwindow``` argument will output messages to a pop-up list.  Double click any item to clear the output.


