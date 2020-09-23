-----------------------------------------------
How to Use Plug-Ins in Your Application
-----------------------------------------------
By Elad Rosenheim

IMPORTANT - In order to run the demo, first register TestPlugIn\TestPlugIn.dll and GreatPlugInPackage\GreatPlugInPackage.dll (you can use regsvr32.exe to do it)

Abstract:

In some scenarios you wish that your app would have the ability to offer more functionality, that would be developed after the app is distributed - without changing and re-distributing the original .exe. You may also wish to offer extra features to specific customers.
The ability to add features to an existing app by third-party vendors is also a major plus - Take Adobe Photoshop or Winamp, for example.
Those little external components that attach and extend your app are most commonly called plug-ins.

My demo shows how to integrate plug-ins easily in a VB application.
It works that way:

--> A Plug-In is implemented as a class in an ActiveX DLL. You give that class a name (well, off course).

--> The host application maintains a list of such class names, and the plug-in inserts its class name to that list when it is installed.

--> When the application is loaded, it loads the list of class names, and from that moment on may use CreateObject to create an instance of the object at run-time, using only the class name (a method known as late-binding). 

--> But hey, how will the app know what methods are available to call?
Simple. The class implements a standard set of functions (you can call it an interface), whose definition is dictated by your app. Your app calls the standard function, and the class does what it knows how to do.

For example:
You have written a paint program. Someone else has written an ActiveX DLL containing a class that can create complex texture effects such as Fractal Textures, given a valid hDC (Windows Device Context Handle) to draw on. The name of the class is FractalTexturizer.PlugInInterface
You, the app creator, has dictated that all plug-ins should implement one function called
DoYourStuff(hDC as long).

When the plug-in is installed, it creates a a new registry entry under your app's registry key, containing the class name, and a "friendly name" for the class that is shown to the end-user.

When the app loads it enumerates all registry entries and reads the class names. For every class name, an entry is added to the "Plug-Ins" app's menu-bar with the friendly name of the plug-in. When the user selects that entry, all you have to do is create an instance of the approperiate class and call DoYourStuff(hDC).

Implementation:
I have written two classes that can help you integrate such a capability:
cPlugIn & cPlugInHandler. Those are generic classes suitable for a wide range of applications.
They assume a standard set of functions that every plug-in should implememnt for most kinds of applications. Those include properties such as Author, Version and ConfigureMyself (Invoking the plug-in's configuration form).

What you are in charge of is thinking what should be the standard interface of functions that a plug-ins should implement for YOUR program - what data should be passed to the plug-in so it will accomplish its task. 

How to inspect the demo: 
1. Don't forget to register the DLLs
2. Run PlugExample.exe - Note that there's nothing under the Plug-Ins menu. The app just contains a text box - and that's that.
3. Run InstallPlugIns.exe from the TestPlugIn sub-dir.
4. When you re-run PlugExample, two plug-ins were added! Check to see what they do to your text - nothing useful really.
5. Run InstallPlugIns.exe from the GreatPlugInPackage sub-dir.
5. Another plug-in is added (re-run the app) - check it out! You must have MS-Word installed to use it.
6. Check out the "List installed Plug-Ins..." option and configure the Spell Checker.

OK, Now that you've seen a demo working, use your force - read the source.

* The prjInstallPlugIns is a small utility that causes plug-ins to add themselves to the app's plug-ins list. It can be used when installing the plug-in with a setup program.

Future Thoughts:

* A problem you may encounter is that "serious" applications store their data in custom classes, and the plug-in DLL should also have a refernce to those class definitions. The solution to this is to put all the shared Class Modules in a separate DLL, that both your app and the plug-in can reference to.

* Moreover, you'll probably want plug-ins to have greater control of your app. It's up to you to define the mechanism - I just laid out the basic infrastructure.

More questions? Ideas?
mailto:eladro@barak-online.net
