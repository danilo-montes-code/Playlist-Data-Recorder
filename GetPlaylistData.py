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


def record_data_on_sheet(main_playlist, other_playlists, sheet):
    for row, og_track_data in enumerate(main_playlist['songs']):  # for every track
        for col, playlist in enumerate(other_playlists):  # for every playlist in other playlists
            check_playlist_for_track(og_track_data['track'], playlist, sheet, row+2, col+2)


def create_sub_playlist_list(playlist):
    results = sp.user_playlist(username, playlist['id'], fields="tracks,next")
    tracks = results['tracks']
    temp_playlist_tracks_list = []
    temp_playlist_tracks_list = add_tracks_to_list(tracks, temp_playlist_tracks_list)

    playlist_dict = {'name': playlist['name'],
                     'number_of_songs': playlist['tracks']['total'], 'songs': []}

    for i in range(0, playlist_dict['number_of_songs']):
        playlist_dict['songs'].append(temp_playlist_tracks_list[i])
    return playlist_dict


def add_tracks_to_list(tracks, playlist):
    for track in tracks['items']:
        playlist.append(track)

    while tracks['next']:
        tracks = sp.next(tracks)
        add_tracks_to_list(tracks, playlist)
    return playlist


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


def get_data():
    # gets the playlists from the user
    playlists = sp.user_playlists(username)

    # gets the main playlist that has all the songs and adds the songs into a list
    spot_main_playlist = sp.user_playlist(username, '1WKZ1xpg8BnmmPgTDDCrI4')
    main_playlist = {'name': spot_main_playlist['name'],
                     'number_of_songs': spot_main_playlist['tracks']['total'], 'songs': []}
    # 2Oi22cH7pgo7AKfGUHih52 - J, 07Rrpr2pjNw4SCyqtPIrqj - D, 1WKZ1xpg8BnmmPgTDDCrI4 - Monarchy
    results = sp.user_playlist(username, spot_main_playlist['id'], fields="tracks,next")
    tracks = results['tracks']
    all_songs_temp = []
    all_songs_temp = add_tracks_to_list(tracks, all_songs_temp)
    for i in range(0, main_playlist['number_of_songs']):
        main_playlist['songs'].append(all_songs_temp[i])

    # gets every other playlist
    sub_playlists = []
    spotify_sub_playlists = get_sub_playlists(playlists, spot_main_playlist)
    for playlist in spotify_sub_playlists:
        sub_playlists.append(create_sub_playlist_list(playlist))

    # opens the excel file and goes to first sheet
    book = openpyxl.load_workbook('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    sheet = book.get_sheet_by_name('Songs')

    # creates the header row
    sheet.cell(1, 1).value = main_playlist['name']
    for i, playlist in enumerate(sub_playlists):
        sheet.cell(1, i + 2).value = playlist['name']

    # puts the songs in the first column
    last_row = 0
    for i, song in enumerate(main_playlist['songs']):
        sheet.cell(i + 2, 1).value = song['track']['name']
        last_row = i + 2

    # saves the excel file's contents
    book.save('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')

    # sorts the song titles in alphabetical order
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    ws = wb.Worksheets('Songs')
    # TODO ws.Range(ws.Cells(2, 1), ws.Cells(last_row, 1)).Sort(Key1=ws.Range('A2'), Order1=1, Orientation=2)

    # Auto resizes the columns, centers the text, and bolds and yellows the top row
    ws.Columns.AutoFit()
    ws.Columns.Style.HorizontalAlignment = -4108
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.color = rgb_to_hex((255, 255, 0))

    # sets the borders to make the sheet easier to read
    ws.Rows(1).Borders.LineStyle = 1
    ws.Columns.Borders(11).LineStyle = 1
    wb.Save()
    excel.Application.Quit()

    # puts in the data for all the songs and saves the file
    book = openpyxl.load_workbook('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    sheet = book.get_sheet_by_name('Songs')
    record_data_on_sheet(main_playlist, sub_playlists, sheet)
    book.save('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open('C:/Users/Rubikscrafter/Documents/MS Excel/Dad\'s Playlist Songs Data.xlsx')
    ws = wb.Worksheets('Songs')
    ws.Columns.Borders(11).LineStyle = 1
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
    clear_data()
    get_data()
    # testing()


if __name__ == '__main__':
    # sets up the credentials for the spotipy object
    client_credentials_manager = SpotifyClientCredentials(client_id='4cd9cb8ecb73460e83343978be07106a',
                                                          client_secret='1a88d95267464a039b4a11a478982e16')
    sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)

    # the username of the user, taken from spotify uri
    username = '22kl7y3a4dhdzvca75vnxxmky'  # 22kl7y3a4dhdzvca75vnxxmky
    main()
