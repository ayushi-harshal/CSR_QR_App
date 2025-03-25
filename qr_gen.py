import qrcode

url = 'http://localhost:5000/'  # Change to your deployed URL
img = qrcode.make(url)
img.save("csr_qr.png")
