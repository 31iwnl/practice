import os
import cv2
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from docx import Document

# Функция для извлечения изображений из документа .docx
def extract_images_from_docx(docx_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    doc = Document(docx_path)
    image_index = 1

    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_name = os.path.join(output_folder, f"image{image_index}.png")
            with open(img_name, "wb") as f:
                f.write(rel.target_part.blob)
            image_index += 1

# Функция для сравнения двух изображений
def compare_images(img1_path, img2_path, threshold=0.9):
    img1 = cv2.imread(img1_path, 0)
    img2 = cv2.imread(img2_path, 0)
    img2 = cv2.resize(img2, img1.shape[::-1])

    difference = cv2.absdiff(img1, img2)
    similarity = 1 - (np.sum(difference) / (img1.shape[0] * img1.shape[1] * 255))

    return similarity > threshold

# Функция для нахождения координат таблицы на изображении
def find_table_coordinates(image_path):
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 100, 10)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    coordinates = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        aspect_ratio = 15 * w / float(h)
        if (w > 250 or 200 < h > 150) and aspect_ratio > 55:
            coordinates.append((x, y, x + w, y + h))
    return coordinates

# Функция для проверки, содержит ли изображение таблицу
def contains_table(image_path, data_folder):
    for data_img in os.listdir(data_folder):
        data_img_path = os.path.join(data_folder, data_img)
        if compare_images(image_path, data_img_path):
            return True
    return False

# Функция для отображения изображения с прямоугольниками вокруг таблиц
def display_image_with_rectangles(image_path, rectangles):
    img = cv2.imread(image_path)
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    for (x1, y1, x2, y2) in rectangles:
        cv2.rectangle(img_rgb, (x1, y1), (x2, y2), (255, 0, 0), 2)

    plt.imshow(img_rgb)
    plt.axis('off')
    plt.show()

def main():
    docx_path = "docs/test/test.docx"
    output_folder = "extracted_images"
    data_folder = "data"

    extract_images_from_docx(docx_path, output_folder)

    for img in os.listdir(output_folder):
        img_path = os.path.join(output_folder, img)
        if contains_table(img_path, data_folder):
            coordinates = find_table_coordinates(img_path)
            print(f"Изображение {img} содержит таблицу. Координаты таблицы: {coordinates}")
            display_image_with_rectangles(img_path, coordinates)

if __name__ == "__main__":
    main()
