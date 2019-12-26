import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
import playlist
import excel_sheet_handling as esh

# Global variables that will be used throughout the entire program
main_playlist = None
other_playlists = []






def main():




if __name__ == '__main__':
    # Sign in to get username?
    client_credentials_manager = SpotifyClientCredentials(client_id='id',
                                                          client_secret='secret')
    sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)
    username = input('What is your spotify uri? '
                 '(right click on your profile name when viewing your profile, hover over share, '
                 'and click the last option)\n')[13:]
    main_playlist_data = input('What is the name of your main playlist (case sensitive)\n')
    main_playlist = playlist.Playlist(username, main_playlist_data)
    other_playlists = []
    main()
