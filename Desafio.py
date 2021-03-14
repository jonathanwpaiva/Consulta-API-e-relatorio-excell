import requests
import xlsxwriter

file = xlsxwriter.Workbook("Relatorio.xlsx")
table = file.add_worksheet()

table.set_column('A:C', 25)
table.set_column('E:G', 25)

title_format = file.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'
})

data_format = file.add_format({
    'border': 1,
    'fg_color': 'yellow'
})

table.merge_range('A1:C1', 'Top Musicas', title_format)
table.write(1,0, "Nome", title_format)
table.write(1,1, "Duracao em segundos", title_format)
table.write(1,2, "Album", title_format)

table.merge_range('E1:G1', 'Álbum', title_format)
table.write(1,4, "Nome", title_format)
table.write(1,5, "Data de lançamento", title_format)
table.write(1,6, "Tipo", title_format)
    
print('----- Consultando deezer api -----\n')

request = requests.get('https://api.deezer.com/artist/3424541')
request2 = requests.get('https://api.deezer.com/artist/3424541/albums')
request3 = requests.get('https://api.deezer.com/album/88022272')

artista = request.json()
print('Informações do Artista')
print('-------------------')
print('Nome: {}\n' .format(artista['name']))
print('Numero de albums: {}\n' .format(artista['nb_album']))
albums = request2.json()
print('Lista dos albums: ')
for x in albums["data"]:    
    print ('', x['title'])
linhaA = 2
for x in albums["data"]:     
    table.write (linhaA,4, x['title'], data_format)
    table.write (linhaA,5, x['release_date'], data_format)
    table.write (linhaA,6, x['record_type'], data_format)
    linhaA += 1
print('Link para o deezer: {}\n' .format(artista['link']))
musicasL = format(artista['tracklist'])
topmusicas = requests.get(musicasL)
musicas = topmusicas.json()
print('Top musics: ')
linhaM = 2
for x in musicas["data"]:    
    print (' Musica: {}'.format(x['title']),'| Duracao: {}'.format(x['duration']),'segundos' ,'| Album: {}' .format(x['album']['title']))    
    table.write (linhaM,0, x['title'], data_format)
    table.write (linhaM,1, x['duration'], data_format)
    table.write (linhaM,2, x['album']['title'], data_format)
    linhaM += 1

print('\nInformações de um album específico')
print('-------------------')
albumE = request3.json()
print('Titulo do album: {}\n' .format(albumE['title']))
print('Data de lançamento: {}\n' .format(albumE['release_date']))
print('Nome da gravadora: {}\n' .format(albumE['label']))
print('Link para o deezer: {}\n' .format(albumE['link']))

musicas2 = format(albumE['tracklist'])
musicasalbum = requests.get(musicas2)
musicl = musicasalbum.json()
print('Lista de musicas do album: ')
for x in musicl["data"]:    
    print ('', x['title'])


file.close()
