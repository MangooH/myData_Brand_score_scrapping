from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import requests
from PIL import Image
from io import BytesIO
import pytesseract
import pandas as pd


class BrandScore:
    def __init__(self):
        try:
            from webdriver_manager import chrome

            print("ChromeDriverManager is installed.")
        except ImportError:
            print(
                "ChromeDriverManager is not installed. Installing ChromeDriverManager..."
            )
            ChromeDriverManager().install()
        self.br = webdriver.Chrome()
        self.iterable_sites = []

    def join_site(self):
        url = "https://brikorea.com/robots.txt"
        self.br.get(url)

        # 텍스트 추출 및 요청 가능 확인
        if "Allow:/" in self.br.find_element(By.CSS_SELECTOR, "pre").text:
            self.br.get(url.split("robots")[0])
        else:
            self.br.get(url)
            print("해당 사이트의 robots.txt 가 변경되었습니다. \n BrandScore class 이용 불가.")

    def search_with_title(self, title):
        search_bar = self.br.find_element(By.ID, "sch_stx")
        search_bar.click()
        search_bar.send_keys(title)
        search_bar.send_keys(Keys.ENTER)

        self.iterable_sites = self.br.find_elements(By.CLASS_NAME, "sch_res_title")
        print(f"해당 검색 결과 {len(self.iterable_sites)}개의 브랜드평판지수 결과가 존재합니다. >")
        for e in self.iterable_sites:
            print(e.text)
        print("extract_df_with_idx(idx) method 를 이용해 월별 데이터를 추출하세요.")

    def extract_df_with_idx(self, idx):
        print("데이터 추출을 시작합니다. 창이 원래대로 돌아올 때 까지 기다려주세요....")
        self.br.find_elements(By.CLASS_NAME, "sch_res_title")[idx].click()
        main_window_handle = self.br.current_window_handle
        # 새탭으로 이동
        for handle in self.br.window_handles:
            if handle != main_window_handle:
                self.br.switch_to.window(handle)
                print(f"{self.br.current_url}로 사이트 이동")
                break
        df = self._extract_score()
        self.br.close()
        self.br.switch_to.window(main_window_handle)
        print("추출 완료")
        return df

    def extract_df_all(self):
        pass

    def _extract_score(self):
        # 회사명 추출 (순위대로)
        elements = self.br.find_elements(
            By.CSS_SELECTOR, ".se-fs-fs16.se-ff-system.se-style-unset"
        )
        if elements:
            corporate = elements[-1].text.split()[7:-2]
            for i, e in enumerate(corporate):
                corporate[i] = e.replace(",", "")
            # 브랜드 지수 이미지 -> text 추출
            url_img = self.br.find_elements(
                By.CSS_SELECTOR, "#bo_v_con > p > span > img"
            )[1].get_attribute("src")
            # 이미지 다운로드
            response = requests.get(url_img)
            img = Image.open(BytesIO(response.content))

            # 이미지를 그레이스케일로 변환
            img_gray = img.convert("L")

            # 이미지에서 텍스트 추출
            text = pytesseract.image_to_string(img_gray, lang="kor+eng")

            # 필요한 텍스트만 추출
            split_text = text.split("\n")
            refine_text = []
            for i, s_t in enumerate(split_text[3:]):
                s_t_split = s_t.split()
                refine_text.append(s_t_split[-5:])

            # 데이터 프레임 생성
            columns = ["금융 업계 이름", "참여지수", "미디어지수", "소통지수", "커뮤니티지수", "브랜드평판지수"]
            refine_dict = {"금융 업계 이름": corporate}
            for i, e in enumerate(zip(*refine_text[: len(corporate)])):
                refine_dict[columns[i + 1]] = list(e)
            df = pd.DataFrame(refine_dict, columns=columns)

            return self._refine_score_data(df)
        else:
            print("주어진 선택자로 찾은 엘리먼트가 없습니다.")
            return False

    def _refine_score_data(self, df):
        df["참여지수"] = pd.to_numeric(
            df["참여지수"].str.replace("[^0-9]", "", regex=True), errors="coerce"
        )
        df["참여지수"] = df["참여지수"].astype(int)
        df["미디어지수"] = pd.to_numeric(
            df["미디어지수"].str.replace("[^0-9]", "", regex=True), errors="coerce"
        )
        df["미디어지수"] = df["미디어지수"].astype(int)
        df["소통지수"] = pd.to_numeric(
            df["소통지수"].str.replace("[^0-9]", "", regex=True), errors="coerce"
        )
        df["소통지수"] = df["소통지수"].astype(int)
        df["커뮤니티지수"] = pd.to_numeric(
            df["커뮤니티지수"].str.replace("[^0-9]", "", regex=True), errors="coerce"
        )
        df["커뮤니티지수"] = df["커뮤니티지수"].astype(int)
        df["브랜드평판지수"] = pd.to_numeric(
            df["브랜드평판지수"].str.replace("[^0-9]", "", regex=True), errors="coerce"
        )
        df["브랜드평판지수"] = df["브랜드평판지수"].astype(int)

        df = df[
            (
                df["참여지수"] + df["미디어지수"] + df["소통지수"] + df["커뮤니티지수"]
                <= df["브랜드평판지수"] * 1.05
            )
            | (
                df["참여지수"] + df["미디어지수"] + df["소통지수"] + df["커뮤니티지수"]
                >= df["브랜드평판지수"] * 0.95
            )
        ]

        return df
