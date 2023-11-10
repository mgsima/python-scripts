import os
import glob
import textwrap
import pathlib as Path
from PIL import Image, ImageDraw, ImageFont

def create_contact_sheet(path, output_path):
    columns=3
    rows=4
    img_width=400
    img_extensions = ['jpg', 'png', 'jpeg', 'gif', 'bmp', 'tiff']
    images = [f for f in glob.glob(os.path.join(path, "*")) if f.split('.')[-1].lower() in img_extensions]
    font = ImageFont.truetype(r"Courier_New.ttf", 40)

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    sheet_counter = 0
    margin = 80 # increased margin
    spacing = 80 # increased spacing
    for i in range(0, len(images), columns * rows):
        sheet_counter += 1
        img_list = images[i:i + columns * rows]

        sheet_width = columns * (img_width + spacing) + margin
        sheet_height = rows * (img_width + spacing) + margin # assuming square images for simplicity
        sheet = Image.new('RGB', (sheet_width, sheet_height), (255, 255, 255)) 
        draw = ImageDraw.Draw(sheet)

        x = y = margin # starting position

        for img_path in img_list:
            img = Image.open(img_path)
            # maintain aspect ratio
            aspect = img.size[0] / img.size[1]
            img_height = int(img_width / aspect)
            img = img.resize((img_width, img_height))

            sheet.paste(img, (x, y))

            # draw the filename below the image
            img_name = os.path.basename(img_path)
            lines = textwrap.wrap(img_name, width=18)  # wrap the text to fit within image width
            y_text = y + img_height
            line_spacing = 10 # additional space between lines in pixels

            for line in lines:
                bbox = draw.textbbox((x, y_text), line, font=font)
                width = bbox[2] - bbox[0]
                height = bbox[3] - bbox[1]
                text_x = x + (img_width - width) / 2
                draw.text((text_x, y_text), line, font=font, fill='black')
                y_text += height + line_spacing



            # move the position for the next image
            x += img_width + spacing
            if x + img_width + margin > sheet_width:
                x = margin
                y += img_height + spacing + 90 # leave room for text

        output_file_name = os.path.join(output_path, f"Endoskopiebilder_Blatt_{sheet_counter}.png")
        sheet.save(output_file_name)

    return sheet_counter


def main():
    path_in = input("Ordner, in dem sich die Fotos befinden: ")
    path_out = input("Ordner, in dem das Ergebnis gespeichert werden soll: ")

    # Uso de la función
    message = create_contact_sheet(Path.Path(path_in), Path.Path(path_out))

    print(f"{message} Kontaktbögen erstellt.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")
