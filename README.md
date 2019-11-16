# DesktopAgent2

The Web keep progressing and each day there are more an more things you can do from a browser.

However sometimes you need to access things that are only on the machine were your browser is running. For example opening local office apps or accessing local devices.

Sometime ago I had provided a similar project called DesktopAgent. That project works fine but there were some lessons learned from it:

* That project was meant as a too general approach.
That project had a concept of plugins for extending functionality and I think that is too complicated. So this project is provided more like a template that provides just the basics, and then you are free 
to adjust it to your needs.

I also switched for a more general communication approach.

* Web Request can be too slow and not that reliable

When I wrote the old desktop agent, IE6 was still around (it is still around I know but let's say its dead), and now all browsers support WebSockets.
So I decided to switch for an implementation based on websockets.

Ok, all those lessons learned bring then DesktopAgent2.

Desktop Agent2 is a simple template to help you get started on creating your own DesktopAgent. 
DesktopAgent2 targets .NET Framework 4.5 