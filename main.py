import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.offsetbox import OffsetImage, AnnotationBbox

DATA_FILENAME = 'test.xlsx'
TOURNAMENT_NAME = '2023年度日本オリエンテーリング選手権大会'
PAPER_SIZE = (11.69, 8.27)  # A4 Size in inches
FONT_NAME = 'Meiryo'

try:
    df = pd.read_csv('データ/data_reformatted.csv')
except FileNotFoundError:
    df = pd.read_excel(f'データ/{DATA_FILENAME}')
    df.to_csv('データ/data_reformatted.csv', index=False, header=True)
    df = pd.read_csv('データ/data_reformatted.csv')

number_kwargs = dict(ha='center', va='center', fontsize=280, color='w', fontname=FONT_NAME)
letter_kwargs = dict(ha='center', va='center', fontsize=35, color='k', fontname=FONT_NAME)
time_kwargs = dict(ha='center', va='center', fontsize=25, color='k', fontname=FONT_NAME)
name_kwargs = dict(ha='center', va='center', fontsize=30, color='k', fontname=FONT_NAME)
furigana_kwargs = dict(ha='center', va='center', fontsize=20, color='k', fontname=FONT_NAME)
university_kwargs = dict(ha='center', va='center', fontsize=30, color='k', fontname=FONT_NAME)


def main() -> None:
    for i in range(len(df['競技者番号'])):
        fig, ax = plt.subplots(figsize=PAPER_SIZE, layout='constrained')

        plt.axis('off')

        plt.text(0.05, 0.9, TOURNAMENT_NAME, fontsize=28, color='k', fontname=FONT_NAME)
        plt.text(0.18, 0.83, df['所属'][i], **university_kwargs)
        plt.text(0.18, 0.76, df['ふりがな'][i], **furigana_kwargs)
        plt.text(0.18, 0.7, df['氏名'][i], **name_kwargs)
        plt.text(0.18, 0.62, df['競技者番号'][i][:5], **letter_kwargs)
        plt.text(0.18, 0.55, df['出走時間'][i][:5], **time_kwargs)
        plt.text(0.5, 0.2, df['競技者番号'][i][-4:], **number_kwargs)

        r = patches.Rectangle(xy=(0.02, 0.), width=0.96, height=0.5, color='brown')
        ax.add_patch(r)

        logo1 = plt.imread(f"所属/{df['所属'][i]}.png")
        imagebox1 = OffsetImage(logo1, zoom=0.2)
        ab1 = AnnotationBbox(imagebox1, (0.9, 0.8), frameon=False)
        ax.add_artist(ab1)

        logo2 = plt.imread('スポンサー/JOA.png')
        imagebox2 = OffsetImage(logo2, zoom=0.9)
        ab2 = AnnotationBbox(imagebox2, (0.6, 0.8), frameon=False)
        ax.add_artist(ab2)

        logo3 = plt.imread(f"国籍/{df['国籍'][i]}.png")
        imagebox3 = OffsetImage(logo3, zoom=0.3)
        ab3 = AnnotationBbox(imagebox3, (0.5, 0.65), frameon=False)
        ax.add_artist(ab3)

        logo4 = plt.imread('スポンサー/SYM.png')
        imagebox4 = OffsetImage(logo4, zoom=0.08)
        ab4 = AnnotationBbox(imagebox4, (0.65, 0.65), frameon=True)
        ax.add_artist(ab4)

        logo5 = plt.imread('スポンサー/OOA.png')
        imagebox5 = OffsetImage(logo5, zoom=0.25)
        ab5 = AnnotationBbox(imagebox5, (0.8, 0.65), frameon=False)
        ax.add_artist(ab5)

        plt.savefig(f"ゼッケン/{df['競技者番号'][i][-4:]}.png", bbox_inches='tight', pad_inches=0, dpi=200)
        plt.close()


if __name__ == '__main__':
    print(df)
    main()
