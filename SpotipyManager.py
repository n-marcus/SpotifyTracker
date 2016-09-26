import spotipy

sp = spotipy.Spotify()


def searchSong(song,artist):
    print("Searching for " + song + artist)
    results = sp.search(q=song + " " + artist , limit=1)
    for i, t in enumerate(results['tracks']['items']):
        print (' ', i, t['name'])
