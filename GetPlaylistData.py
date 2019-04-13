import openpyxl
import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
import win32com.client


'''
Checks a given playlist for a track.

Args:
    track: id of the track to search for
    playlist: playlist dictionary that will be searched for the track
    sheet: the excel sheet to write data for finding a song onto
    row: row that the song is in on the excel sheet
    col: column that the playlist is in on the excel sheet
'''
def check_playlist_for_track(track, playlist, sheet, row, col):
    for playlist_song in playlist['songs']:
        if playlist_song['track']['id'] == track['id']:
            print(f'Found song {track["name"]} in {playlist["name"]}')
            sheet.cell(row, col).value = 'X'
            break


'''
Loops to write the data on the excel sheet

Args:
    main_playlist: dictionary of the playlist that the songs are being are taken from
    other_playlists: list of dictionaries of playlists that are being searched for the songs
    sheet: excel sheet to write the data on
    index: pushes the row counter forward as the song number changes, as each search is done in sets of 100
'''
def record_data_on_sheet(main_playlist, other_playlists, sheet, index):
    for row, og_track_data in enumerate(main_playlist['songs']):  # for every track
        print()
        print(f'Searching for {og_track_data["track"]["name"]} in playlists.')
        for col, playlist in enumerate(other_playlists):  # for every playlist in other playlists
            print(f'Checking playlist: {playlist["name"]}')
            check_playlist_for_track(og_track_data['track'], playlist, sheet, index+row+2, 49+col+3)


'''
Appends to the list containing the dictionaries of all the playlists besides the main one

Args:
    playlist: the playlist to be appended to the list
    
Return:
    playlist dictionary after adding its tracks to said dictionary
'''
def create_sub_playlist_list(playlist):
    playlist_dict = {'name': playlist['name'],
                     'id': playlist['id'],
                     'number_of_songs': playlist['tracks']['total'],
                     'songs': []}

    results = sp.user_playlist(username, playlist['id'], fields="tracks,next")
    tracks = results['tracks']
    return add_tracks_to_list(tracks, playlist_dict)


'''
Adds tracks to the dictionary of the playlist that is passed in

Args:
    tracks: the tracks to be added to the dictionary
    playlist_dict: the playlist dictionary to add the songs to
    index: for the main playlist, shifts the songs down to the correct set of 100 to store
    
Return:
    playlist_dict: the created playlist dictionary
'''
def add_tracks_to_list(tracks, playlist_dict, index=0):
    if playlist_dict['name'] != 'All My Songs':
        for track in tracks['items']:
            playlist_dict['songs'].append(track)
        while tracks['next']:
            tracks = sp.next(tracks)
            add_tracks_to_list(tracks, playlist_dict)

    if playlist_dict['name'] == 'All My Songs':
        shift = int(index/100)
        if tracks['next']:
            for i in range(0, shift):
                tracks = sp.next(tracks)
            for track in tracks['items']:
                playlist_dict['songs'].append(track)
    return playlist_dict


'''
Gets the other public user playlists that are not the main playlist

Args:
    playlists: all of the user's public playlists
    main_playlist: the dictionary of the main playlist
  
Return:
    sub_playlists: a list of all the user's public playlists that are not the main playlist
'''
def get_sub_playlists(playlists, main_playlist):
    sub_playlists = []
    for playlist in playlists['items']:
        if playlist['owner']['id'] == username and playlist['id'] != main_playlist['id']:
            sub_playlists.append(playlist)
    return sub_playlists


'''
Converts a passed in rgb tuple to hexadecimal, useful for changing the background of the top row on the excel sheet, making the column headers more distinguishable from other data

Args:
    rgb: a tuple a rgb values corresponding to a color
    
Return:
    the hexadecmial value of the passed in rgb value
'''
def rgb_to_hex(rgb):
    bgr = (rgb[2], rgb[1], rgb[0])
    str_value = '%02x%02x%02x' % bgr
    return int(str_value, 16)


'''
Gets the data and writes it to the excel sheet; essentially the main method

Args:
    playlist_name: id of the main playlist
    index: buffer number to shift forward the obtained sets of 100 songs from the main playlist
    header_already_made: a boolean variable that is true if the header row on the excel sheet was already made and false otherwise
'''
def get_data(playlist_name, index, header_already_made):
    # gets the playlists from the user
    # playlists = sp.user_playlists(username)
    playlists2 = sp.user_playlists(username, offset=50)

    # gets the main playlist that has all the songs and adds the songs into a dictionary
    results = sp.user_playlist(username, playlist_name, fields="tracks,next")
    main_playlist = {'name': sp.user_playlist(username, playlist_name)['name'],
                     'id': sp.user_playlist(username, playlist_name)['id'],
                     'number_of_songs': sp.user_playlist(username, playlist_name)['tracks']['total'],
                     'songs': []}
    tracks = results['tracks']
    main_playlist = add_tracks_to_list(tracks, main_playlist, index)

    # gets every other playlist and makes dictionaries for them
    sub_playlists = []
    all_subs = [playlists2]
    for playlist_set in all_subs:
        spotify_sub_playlists = get_sub_playlists(playlist_set, main_playlist)
        for playlist in spotify_sub_playlists:
            sub_playlists.append(create_sub_playlist_list(playlist))

    # opens the excel file and goes to first sheet
    book = openpyxl.load_workbook('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    sheet = book.get_sheet_by_name('Songs')

    # only makes the header row if it wasn't already made
    if not header_already_made:
        # creates the header row
        sheet.cell(1, 1).value = main_playlist['name']
        sheet.cell(1, 2).value = "Artist(s)"
        for i, playlist in enumerate(sub_playlists):
            sheet.cell(1, i + 3).value = playlist['name']

    # puts the songs in the first column
    temp_artists = ''
    multiple_artists = False
    for i, song in enumerate(main_playlist['songs']):
        sheet.cell(index + i + 2, 1).value = song['track']['name']
        for artist in song['track']['artists']:
            if multiple_artists:
                temp_artists += ', '+artist['name']
            else:
                temp_artists += artist['name']
            multiple_artists = True
        sheet.cell(index + i + 2, 2).value = temp_artists
        temp_artists = ''
        multiple_artists = False

    # saves the excel file's contents
    book.save('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')

    # puts in the data for all the songs and saves the file
    book = openpyxl.load_workbook('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    sheet = book.get_sheet_by_name('Songs')
    record_data_on_sheet(main_playlist, sub_playlists, sheet, index)
    book.save('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    format_cells()


def format_cells():
    # opens sheet with win32 for formatting
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    ws = wb.Worksheets('Songs')
    ws.Columns.Borders(11).LineStyle = 1

    # Auto resizes the columns, centers the text, and bolds and yellows the top row
    ws.Columns.AutoFit()
    ws.Columns.Style.HorizontalAlignment = -4108
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.color = rgb_to_hex((255, 255, 0))

    # sets the borders to make the sheet easier to read
    ws.Rows(1).Borders.LineStyle = 1
    ws.Columns.Borders(11).LineStyle = 1
    # ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).Sort(Key1=ws.Range(ws.Cells(2, 1), ws.Cells(last_row, 1)),
    # Order1=1, Orientation=2)

    # saves and quits
    wb.Save()
    excel.Application.Quit()


def main():
    print              
                  
    for i in range(600, 3201, 100):
        index = i  # Should be the row number of the last entered data -1
        header_already_made = True
        get_data('2Oi22cH7pgo7AKfGUHih52', index, header_already_made)

    # 2Oi22cH7pgo7AKfGUHih52 - J, 07Rrpr2pjNw4SCyqtPIrqj - D, 1WKZ1xpg8BnmmPgTDDCrI4 - Monarchy


if __name__ == '__main__':
    # sets up the credentials for the spotipy object
    client_credentials_manager = SpotifyClientCredentials(client_id='4cd9cb8ecb73460e83343978be07106a',
                                                          client_secret='1a88d95267464a039b4a11a478982e16')
    sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)
    
    # sets up the variables that will be used throughout the script
    username = ''
    main_playlist = {}
    sub_playlists = []
    sheet = None
    
    # asks the user for their username (hopefully temporary), taken from spotify uri
    user = input('What is your spotify api? (right click on your profile name when viewing your profile and click the last option)')
    username = user[13:]
    main()
