# DesktopAgent2

The Web keeps progressing and each day there are more things you can do from a browser.

However, sometimes you need to access things that are only on the machine where your browser is running. For example, opening local office apps or accessing local devices.

![BeforeDesktopAgent](Images\BeforeDesktopAgent2.png)

Some time ago I had provided a similar project called [DesktopAgent](https://github.com/orellabac/WebMap.DesktopAgent). That project works fine but there were some lessons learned from it:

* That project was meant as a too general approach.
That project had a concept of plugins for extending functionality and I think that is too complicated. So this project is provided more like a template that provides just the basics, and then you are free 
to adjust it to your needs.

I also switched for a more general communication approach.

* Web Request can be too slow and not that reliable

When I wrote the old desktop agent, IE6 was still around (it is still around I know but let's say its dead), and now all browsers support WebSockets.
So I decided to switch for an implementation based on WebSockets.

Ok, all those learned lessons bring then DesktopAgent2.

![AfterDesktopAgent](Images\AfterDesktopAgent2.png)

Desktop Agent2 is a simple template to help you get started on creating your DesktopAgent. 
DesktopAgent2 targets .NET Framework 4.5, so it can be used even from Windows 7.

Building
========

To build the code just open `DesktopAgent2.csproj` in Visual Studio

Testing
=======

If you just want to test *DesktopAgent2* you can build it from source of download the binary from the Releases link.

First you need to run the Agent ![Desktop Agent](Images\DesktopAgent2.png)

Next open the `test.html' page.
![TestPage](Images\testpage.png)

When you press connect, the page will send a connection request to the agent. Once connected the agent will indicate that it has a new client

![One client](Images\AgentWithClient.png)

Now you can press the test buttons. For example press `Start Word`

The Agent will display the json message it received 

```json
{"action":"Word","command":"open"}
```

In general the Agent usage is pretty simple:

```js
       // create a JSON
       var msg = {'action':'Excel','command':'open'};
       //Send the JSON message
       agent.Send(JSON.stringify(msg)).
       // react to the response
       then(function() {
            alert('excel was opened');
        });
```

On the Agent you just modify the code to insert your custom actions:

```C#
    private void ExecuteActions(IWebSocketConnection webSocket, string message)
        {
            Invoke(new Action(() => txtLog.AppendText(message)));

            try
            {
                var action = JObject.Parse(message);
                var command = action["action"].Value<string>();
                switch (command)
                {
                    //Add your actions in this switch
                    case "WSH":
                        this.RunWSH(webSocket, action);
                        break;
                    case "ExecProgram":
                        this.RunProgram(webSocket, action);
                        break;
                    case "Excel":
                        this.ExcelAction(webSocket, action);
                        break;
                    case "Word":
                        this.WordAction(webSocket, action);
                        break;
                }
            }
            catch
            {
                // Just ignore
            }
        }
```

Secure communication
=====================

Enabling secure connections requires two things: using the scheme wss instead of ws, and pointing to an x509 certificate containing public and private key

The changes in code are easy. Just two lines:

```C#
    var server = new WebSocketServer("wss://0.0.0.0:996");
    server.Certificate = new System.Security.Cryptography.X509Certificates.X509Certificate2(@"M:\Program Files\OpenSSL-Win64\bin\secondtest.pfx", "test");
```

To create a certificate you can follow this guide: https://github.com/statianzo/Fleck/issues/214#issuecomment-364413879


