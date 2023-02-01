#include ..\Lib\UIA.ahk
/**
 * This is a small library to demonstrate how UIA could be used to automate the Spotify app.
 * It has not been tested enough to be actually used.
 */

; Don't bring Spotify to focus automatically
UIA.AutoSetFocus := False

F1::Spotify.TogglePlay()
F11::Spotify.ToggleFullscreen()
^d::Spotify.Toast("Playing song: " (song := Spotify.CurrentSong).Name "`nArtist: " song.Artist "`nPlay time: " song.Time " / " song.Length)
^l::
{
    song := Spotify.CurrentSong
    Spotify.Toast("You " (!Spotify.LikeState ? "liked " : "removed a like from ") song.Name " by " song.Artist)
    Spotify.ToggleLike()
}
^Left::Spotify.NextSong()
^Right::Spotify.PreviousSong()
^+::Spotify.Volume += 10
^-::Spotify.Volume -= 10
^m::Spotify.ToggleMute()

class Spotify {
    static winExe := "ahk_exe Spotify.exe"
    static winTitle => WinGetTitle(this.winExe)
    static exePath := A_AppData "\Spotify\Spotify.exe"

    ; Internal method to show a Toast message, but before that remove the previous one
    static Toast(message) {
        TrayTip
        A_IconHidden := true
        Sleep 200
        A_IconHidden := false
        TrayTip(message, "Spotify info")
    }
    ; Internal methods to get some commonly used Spotify UIA elements
    static GetSpotifyElement() => UIA.ElementFromHandle(Spotify.winExe)[1]
    static GetLikeElement() => Spotify.GetSpotifyElement().ElementFromPath({T:26, i:3}, {T:26,N:"Now playing: ",mm:2}, {T:0,N:"Library", mm:2})
    static GetCurrentSongElement() => Spotify.FullscreenState ? Spotify.GetSpotifyElement() : Spotify.GetSpotifyElement()[{T:26, i:3}]

    static LikeState {
        get => Spotify.GetLikeElement().ToggleState
        set => Spotify.GetLikeElement().ToggleState := value
    }
    static Like() => Spotify.LikeState := 1
    static RemoveLike() => Spotify.LikeState := 0
    static ToggleLike() => Spotify.LikeState := !Spotify.LikeState

    static CurrentSong {
        get {
            contentEl := Spotify.GetCurrentSongElement()
            return {
                Name:contentEl[{Type:"Group"},{Type:"Link",i:2}].Name, 
                Artist:contentEl[{Type:"Group"},{Type:"Link",i:3}].Name, 
                Time:contentEl[{Type:"Text", i:1}].Name,
                Length:contentEl[{Type:"Text", i:3}].Name
            }
        }
    }
    static CurrentSongState {
        get {
            try return Spotify.GetCurrentSongElement()[[{Name:"Play"}, {Name:"Pause"}]].Name == "Play"
            catch
                throw Error("Play/Pause button not found!", -1)
        }
        set {
            if value != Spotify.CurrentSongState
                try Spotify.GetCurrentSongElement()[[{Name:"Play"}, {Name:"Pause"}]].Click()
        }
    }
    static Play() => Spotify.CurrentSongState := 1
    static Pause() => Spotify.CurrentSongState := 0
    static TogglePlay() => Spotify.CurrentSongState := !Spotify.CurrentSongState

    static NextSong() => Spotify.GetCurrentSongElement()[{Name:"Next"}].Click()
    static PreviousSong() => Spotify.GetCurrentSongElement()[{Name:"Previous"}].Click()

    static FullscreenState {
        get => Spotify.GetSpotifyElement()[-1].Type == UIA.Type.Button
        set {
            WinActivate(Spotify.winExe)
            WinWaitActive(Spotify.winExe,,1)
            if Spotify.FullscreenState
                Spotify.GetSpotifyElement()[-1].Click()
            else
                Spotify.GetCurrentSongElement()[-1].Click()
        }
    }
    static ToggleFullscreen() => Spotify.FullscreenState := !Spotify.FullscreenState

    static MuteState {
        get => Spotify.GetCurrentSongElement()[{Type:"Button", i:-2}] == "Mute"
        set {
            currentState := Spotify.MuteState
            if Value && !currentState
                Spotify.GetCurrentSongElement()[{Type:"Button", i:-2}].Click()
            if !Value && currentState
                Spotify.GetCurrentSongElement()[{Type:"Button", i:-2}].Click()

        }
    }
    static ToggleMute() => Spotify.MuteState := !Spotify.MuteState
    static Mute() => Spotify.MuteState := 1
    static Unmute() => Spotify.MuteState := 0

    static Volume {
        get => Spotify.GetCurrentSongElement()[{Type:"Slider",i:-1}].Value
        set => Spotify.GetCurrentSongElement()[{Type:"Slider",i:-1}].Value := value
    }
}