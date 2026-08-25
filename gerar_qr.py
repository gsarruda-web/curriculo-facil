import sys
import qrcode
url = sys.argv[1] if len(sys.argv)>1 else 'https://SEU-ENDERECO-AQUI'
img=qrcode.make(url)
img.save('qr_curriculo.png')
print('QR gerado para:',url)
