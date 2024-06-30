import os
import cv2
import numpy as np
from PIL import Image
from docx import Document


# Функция для извлечения изображений из документа .docx
def extract_images_from_docx(docx_path, output_folder):
    # Проверяем, существует ли папка, если нет - создаем
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Открываем документ
    doc = Document(docx_path)
    image_index = 1

    # Проходим по всем связям документа
    for rel in doc.part.rels.values():
        # Если связь является изображением
        if "image" in rel.target_ref:
            # Формируем имя файла для изображения
            img_name = os.path.join(output_folder, f"image{image_index}.png")
            # Сохраняем изображение в файл
            with open(img_name, "wb") as f:
                f.write(rel.target_part.blob)
            image_index += 1


# Функция для сравнения двух изображений
def compare_images(img1_path, img2_path, threshold=0.9):
    # Загружаем изображения в градациях серого
    img1 = cv2.imread(img1_path, 0)
    img2 = cv2.imread(img2_path, 0)

    # Изменяем размер второго изображения для совпадения с первым
    img2 = cv2.resize(img2, img1.shape[::-1])

    # Вычисляем разницу между изображениями
    difference = cv2.absdiff(img1, img2)
    # Вычисляем схожесть изображений
    similarity = 1 - (np.sum(difference) / (img1.shape[0] * img1.shape[1] * 255))

    # Возвращаем True, если схожесть выше порога
    return similarity > threshold


# Функция для нахождения координат таблицы на изображении
def find_table_coordinates(image_path):
    # Загружаем изображение
    img = cv2.imread(image_path)
    # Преобразуем изображение в градации серого
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # Применяем бинарный порог для выделения таблицы
    _, thresh = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV)
    # Ищем контуры на бинаризированном изображении
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Список для хранения координат таблиц
    coordinates = []
    for contour in contours:
        # Вычисляем ограничивающий прямоугольник для каждого контура
        x, y, w, h = cv2.boundingRect(contour)
        # Добавляем координаты в список
        coordinates.append((x, y, x + w, y + h))
    return coordinates


# Функция для проверки, содержит ли изображение таблицу
def contains_table(image_path, data_folder):
    for data_img in os.listdir(data_folder):
        data_img_path = os.path.join(data_folder, data_img)
        # Если изображение из docx похоже на одно из изображений в папке data, возвращаем True
        if compare_images(image_path, data_img_path):
            return True
    return False


def main():
    # Путь к документу .docx
    docx_path = "docs/test/test.docx"
    # Папка для извлеченных изображений
    output_folder = "extracted_images"
    # Папка с изображениями для сравнения
    data_folder = "data"

    # Извлекаем изображения из документа
    extract_images_from_docx(docx_path, output_folder)

    # Проверяем, содержат ли извлеченные изображения таблицы
    for img in os.listdir(output_folder):
        img_path = os.path.join(output_folder, img)
        if contains_table(img_path, data_folder):
            # Находим координаты таблиц на изображении
            coordinates = find_table_coordinates(img_path)
            # Выводим имя изображения и координаты таблиц
            print(f"Изображение {img} содержит таблицу. Координаты таблицы: {coordinates}")


if __name__ == "__main__":
    main()
