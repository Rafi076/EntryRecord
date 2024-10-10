import qrcode
from PIL import Image
import os

def generate_qr_code(data, phone_number):
    # Generate the QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill="black", back_color="white")

    # Save the QR code with the phone number as the filename
    folder = "QR images"
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    img_path = os.path.join(folder, f"{phone_number}.png")
    img.save(img_path)
