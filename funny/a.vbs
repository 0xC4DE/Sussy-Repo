With CreateObject("WMPlayer.OCX")
    .url = "http://www.opensus.org/funny/fart-with-reverb.mp3"
    .controls.play
    .settings.volume = 100
    Do
        WScript.Sleep 100
    Loop Until .playState = 1
End With
