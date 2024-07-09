import libtorrent as lt
import time

# Create a session
ses = lt.session()

# Add a torrent
info = lt.parse_magnet_uri('magnet:?xt=urn:btih:5FF2F4CB706C673DFE042922B56B2C7A944EFE2B&dn=Mirzapur+%282024%29+Season+03+Hindi+S03+Complete+1080p+AMZN+WEB-DL+-+5.7+GB+-+DDP+5.1+AVC+x264+Esub+-+SHADOW+%5BProtonMovies%5D&tr=udp%3A%2F%2Ftracker.opentrackr.org%3A1337%2Fannounce&tr=udp%3A%2F%2Fopen.stealth.si%3A80%2Fannounce&tr=udp%3A%2F%2Ftracker.torrent.eu.org%3A451%2Fannounce&tr=udp%3A%2F%2Fexplodie.org%3A6969%2Fannounce&tr=udp%3A%2F%2Fopen.demonii.com%3A1337%2Fannounce&tr=udp%3A%2F%2Ftracker.tiny-vps.com%3A6969%2Fannounce&tr=http%3A%2F%2Fbt.okmp3.ru%3A2710%2Fannounce&tr=https%3A%2F%2Ftracker.tamersunion.org%3A443%2Fannounce&tr=http%3A%2F%2Ftracker.vraphim.com%3A6969%2Fannounce&tr=http%3A%2F%2Ftracker2.dler.org%3A80%2Fannounce&tr=udp%3A%2F%2Fexodus.desync.com%3A6969%2Fannounce&tr=udp%3A%2F%2Ftracker.0x7c0.com%3A6969%2Fannounce&tr=https%3A%2F%2Ftracker.gbitt.info%3A443%2Fannounce&tr=http%3A%2F%2Ftracker.gbitt.info%3A80%2Fannounce&tr=udp%3A%2F%2Ftracker.opentrackr.org%3A1337%2Fannounce&tr=http%3A%2F%2Ftracker.openbittorrent.com%3A80%2Fannounce&tr=udp%3A%2F%2Fopentracker.i2p.rocks%3A6969%2Fannounce&tr=udp%3A%2F%2Ftracker.internetwarriors.net%3A1337%2Fannounce&tr=udp%3A%2F%2Ftracker.leechers-paradise.org%3A6969%2Fannounce&tr=udp%3A%2F%2Fcoppersurfer.tk%3A6969%2Fannounce&tr=udp%3A%2F%2Ftracker.zer0day.to%3A1337%2Fannounce')  # replace with your actual magnet link
h = ses.add_torrent({"ti": lt.torrent_info(info), "save_path": "./"})

print("downloading", h.name())
while not h.is_seed():
    s = h.status()

    print("download rate: ", s.download_rate / 1000, "kB/s")
    print("progress: %.2f%%" % (s.progress * 100))
    print("peers: ", s.num_peers)
    print("")

    time.sleep(1)

print(h.name(), "complete")
