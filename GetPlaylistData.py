import openpyxl
import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
import win32com.client


def check_playlist_for_track(track, playlist, sheet, row, col):
    for playlist_song in playlist['songs']:
        if playlist_song['track']['id'] == track['id']:
            print(f'Found song {track["name"]} in {playlist["name"]}')
            sheet.cell(row, col).value = 'X'
            break


def record_data_on_sheet(main_playlist, other_playlists, sheet, index):
    for row, og_track_data in enumerate(main_playlist['songs']):  # for every track
        print()
        print(f'Searching for {og_track_data["track"]["name"]} in playlists.')
        for col, playlist in enumerate(other_playlists):  # for every playlist in other playlists
            print(f'Checking playlist: {playlist["name"]}')
            check_playlist_for_track(og_track_data['track'], playlist, sheet, index+row+2, col+3)


def create_sub_playlist_list(playlist):
    playlist_dict = {'name': playlist['name'],
                     'id': playlist['id'],
                     'number_of_songs': playlist['tracks']['total'],
                     'songs': []}

    results = sp.user_playlist(username, playlist['id'], fields="tracks,next")
    tracks = results['tracks']
    return add_tracks_to_list(tracks, playlist_dict)


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


def get_sub_playlists(playlists, main_playlist):
    sub_playlists = []
    for playlist in playlists['items']:
        if playlist['owner']['id'] == username and playlist['id'] != main_playlist['id']:
            sub_playlists.append(playlist)
    return sub_playlists


def rgb_to_hex(rgb):
    bgr = (rgb[2], rgb[1], rgb[0])
    str_value = '%02x%02x%02x' % bgr
    # print(strValue)
    i_value = int(str_value, 16)
    return i_value


def get_data(playlist_name, index, header_already_made):
    # gets the playlists from the user
    playlists = sp.user_playlists(username)

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
    spotify_sub_playlists = get_sub_playlists(playlists, main_playlist)
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


def clear_data():
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    ws = wb.Worksheets('Songs')
    ws.Rows.Clear()
    ws.Columns.Clear()
    ws.Rows.Borders(12).LineStyle = -4142
    ws.Columns.Borders(11).LineStyle = -4142
    ws.Columns.ColumnWidth = 10

    wb.Save()
    excel.Application.Quit()


def testing():
    playlists = sp.user_playlists(username)
    for thing in playlists['items']:
        print(thing['name'])


def main():
    # clear_data()
    index = 300  # 300
    header_already_made = True
    get_data('2Oi22cH7pgo7AKfGUHih52', index, header_already_made)
    # 2Oi22cH7pgo7AKfGUHih52 - J, 07Rrpr2pjNw4SCyqtPIrqj - D, 1WKZ1xpg8BnmmPgTDDCrI4 - Monarchy
    # testing()


if __name__ == '__main__':
    # sets up the credentials for the spotipy object
    client_credentials_manager = SpotifyClientCredentials(client_id='4cd9cb8ecb73460e83343978be07106a',
                                                          client_secret='1a88d95267464a039b4a11a478982e16')
    sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)

    # the username of the user, taken from spotify uri
    username = 'revjose49'  # 22kl7y3a4dhdzvca75vnxxmky - D, revjose49 - J
    main()
