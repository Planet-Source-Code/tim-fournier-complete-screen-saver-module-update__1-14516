Project: Screen Saver Module
Author: Tim Fournier
Contact: tim_fournier@hotmail.com

Last Update: February 24, 2002

Something I've noticed about many of the screen savers submitted is that although they are graphically excellent, they do not handle all the events which a Screen Saver normally should (changing passwords, a configure dialog box, etc.). Although some even did this, they maybe had one of these things, and it is poorly commented.

All of this led me to create this Module, along with a couple forms to see how easy it is to make them interact. I tried to make as many of the functions of the Screen Saver itself get handled by this Module, and have heavily commented everything in the Module. Now you can simply take your background and events, and throw them in without any hassle.

This screen saver module runs quite similar to the one that was submitted by David Saunders, Unusal Cars Screen Saver and I suggest you have a look at his to get another view of how this can be accomplished.

UPDATE NOTICE (Feb 2002): This application has been updated to work in the WinNT environment. Special thanks to Kyle Burns for pointing out this problem and providing an excellent solution.

Just a bit of information if you are going to use this module to set up a Screen Saver of your own design:

-You can freely change the project title to whatever you make your screen saver, just make sure you also change it in the properties settings, as the Registry entries are based on the App.Title.

-Any objects you add to frmMain should have a call to the ShutDown procedure in their keypress, mousedown and mousemove events.

-Put any activities in your Screen Saver in the Timer1_Timer event of frmMain. Adjust the Interval accordingly

-If your graphics require the use of any API, you would be better off declaring them Privately in frmMain, as they will easily get lost in those of the module (Somewhere near 20 APIs are declare there). This will help you keep things organized, the Module should be able to handle everything else for you.

Some things I am still working on, and hope to add in a future upload:

-Drawing the Current Desktop Image to the Screen Saver Background, my attempt is in the Sub Procedure GetDesktop.