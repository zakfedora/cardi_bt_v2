from pdf2image import convert_from_path

def convert_to_image(pdf_file):
    images = []

    try:
        # Convert PDF to list of images
        images = convert_from_path(pdf_file)

        for i, image in enumerate(images):
            # You can save the images to files, display them, or perform other operations
            image.save(f'schema_unifulaire_{i + 1}.png')

    except Exception as e:
        print(f"Error converting PDF to images: {e}")

    return 'schema_unifulaire_1.png'