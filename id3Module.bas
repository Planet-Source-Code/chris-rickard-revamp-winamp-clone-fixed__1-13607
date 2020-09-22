Attribute VB_Name = "id3Module"
Public Type Id3                 'This type is standard for
Title As String * 30            ' Id3 Tags
Artist As String * 30           ' Although later versions
Album As String * 30            ' use comments for 28 bytes
sYear  As String * 4            ' and they use the 2 remaining  bytes for "TrackNumber"!
Comments As String * 30
Genre As Byte
End Type

Public id3Info As Id3
' Declare a variable as the id3 type
Public GenreArray() As String         ' we use this array to fill all the Genre's ( look in form load)

Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
' Thank you Leigh Bowers
' We can use the split function to fill this into an array

Public Function GetId3(FileName As String)
Dim TaG As String * 3               ' We use this variable to make sure the file has an ID3TAG
Open FileName For Binary As #1      ' we open the file as binary for total control (we need it for the Genre part)
Get #1, FileLen(FileName) - 127, TaG    ' Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
If TaG = "TAG" Then                     ' "TAG" is put at position filesize-127 to show that this file indeed contains an Id3
Get #1, FileLen(FileName) - 124, id3Info    ' if the file has a tag, we put it into our earlier declared variable id3info
Else
MsgBox "This Mp3 Does Not Contain an ID3 Tag"       ' if the "TAG" wasnt at position filesize-127
End If
Close #1                                            ' close the file

' Now about the Genre
' It works like this, it contains a code in numbers ranging form 1 to 147
' each of these numbers represents a certain Genre like "HipHop" = 7 etc etc.
' the guy who maid the Id3 Tags made a list for the codes and there were originally 80 of them
' then the dudes at winamp extended this so today there are about 150
' this is a pain in the ass to figure out, still there are some info about this on the www.
' Now, a very cool person by the name of Leigh Bowers, has done this. you can search for the code
' on planet source, "MP3Snatch v2.0", but that code has a couple of flaws in the genre part as it uses
' a string*21 instead of a Byte, and on that code you cant write the tag, only read it.
' so i have included Leighs code wich has 147 of the Genre's, very cool.


' if you want the Genre directly, try filling a combobox with the GenreArray and then use combo1.listindex to match the Genre(code) (number)
End Function

Public Function SaveId3(FileName As String, Mp3Info As Id3)

On Error GoTo blokie
GoTo mick
blokie:
MsgBox "There was an unexpected error saving ID3 tag"
Exit Function
mick:
'MsgBox FileLen(Filename)
Dim TaG As String * 3               ' We use this variable to make sure the file has an ID3TAG
Open FileName For Binary As #1      ' we open the file as binary for total control (we need it for the Genre part)
Get #1, FileLen(FileName) - 127, TaG    ' Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
If TaG = "TAG" Then                     ' "TAG" is put at position filesize-127 to show that this file indeed contains an Id3
Put #1, FileLen(FileName) - 124, Mp3Info    ' if the file has a tag, we put our new information in the file
Else
Put #1, FileLen(FileName) - 127, "TAG"      ' else we put the "TAG" there first,
Close #1
Call SaveId3(FileName, Mp3Info)                               ' then we call this function again so we fill the info this time
End If
Close #1                                            ' close the file
' Remember when filling the Mp3info variable
' set the genre part to the number corresponding to the Genre
' it will save it as 1 byte
End Function





