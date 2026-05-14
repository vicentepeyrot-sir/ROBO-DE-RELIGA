import pyautogui as pya
from directories import Directories
import os
from time import sleep
import traceback
from threading import Event


class ImageFinder:
    def __init__(self, _dir: Directories):
        self._dir = _dir
        self.cancel_event = Event()

    def find_image(self, img: str, sec: int, quality=0.8) -> bool:
        try:
            self.cancel_event.clear()
            for _ in range(sec):
                if self.cancel_event.is_set():
                    return False
                img_path = os.path.join(self._dir.dir_images, f"{img}.jpg")
                try:
                    img_coordinates = pya.locateCenterOnScreen(
                        img_path, confidence=quality)
                except:
                    img_coordinates = None
                if img_coordinates is not None:
                    return True
                sleep(1)
            return False
        except FileNotFoundError as e:
            print(e)
            traceback.print_exc()
            raise FileNotFoundError(f'Error locating image: {img_path}')

    def click(self, img, quality=0.8):
        """
        Clicar na imagem que deseja na tela
        :parameter img: imagem que deseja a ser clicada
        """
        img_path = os.path.join(self._dir.dir_images, f"{img}.jpg")
        try:
            img_cordinates = pya.locateCenterOnScreen(
                img_path, confidence=quality)
            if img_cordinates != None:
                pya.leftClick(img_cordinates, duration=0.4)
                return True
            else:
                return False

        except FileNotFoundError as e:
            print(e)
            traceback.print_exc()
            raise FileNotFoundError(f'Error locating image: {img_path}')

    def cancel(self):
        self.cancel_event.set()
