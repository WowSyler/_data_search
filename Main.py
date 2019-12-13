import FilmList
import FilmListDetay


flist = FilmList.FilmListXml()
flist.SaveList()

flistDetail = FilmListDetay.Crawler()
flistDetail.WriteDetail()
