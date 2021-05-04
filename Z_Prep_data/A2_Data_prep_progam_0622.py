import pandas as pd
import sys

sys.path.append('C:/GDSRC/K_main_mymodules')
from K_main_mymodules import Fa_pre,Fa_rwb
Fa_pre.status_check()
# --------------------------------------------------Loading DATA
df = Fa_rwb.Loader_L()
j_df = df
# --------------------------------------------------
j_df = j_df.drop(['차트번호','검사코드','접수일자/접수번호'], axis=1)
st_item = j_df['검사명'].tolist()
st_len = (len(st_item))
print('st_len  ====>    ',st_len)
df1 = j_df

# --------------------------------------------------
def code_1(i,x,y_1):
    for y in x:
        df1.loc[df1['검사명'] == y, 'dt_code_1'] = (y_1 + str(i).zfill(3))
        i = i + 1

code_1(0,['T.Protein', 'Albumin', 'T.Bilirubin',
            'AST(sGOT)', 'ALT(sGPT)', 'ALP', 'GGT(γ-GTP)', 'Uric acid',
            'LDH', 'CK,total(CPK)', 'Cholesterol, total',
            'Triglyceride', 'HDL-cholesterol', 'LDL-cholesterol',
            'S_Glucose', 'HbA1C', 'Amylase', 'Lipase',
            'BUN', 'Creatinine', 'Calcium (Ca)', 'Phosphorus (P)',
            'Sodium (Na)', 'Chloride (Cl)', 'Potassium (K)'],  'A1.')

code_1(27,['WBC count', 'RBC count', 'Hemoglobin',
           'Hematocrit', 'MCV', 'MCH', 'MCHC', 'RDW', 'Ferritin', 'Iron (Fe)', 'TIBC',
           'Folate', 'Platelet', 'MPV', 'PCT', 'PDW',
           'Differential count', 'Lymphocyte', 'Monocyte', 'Neutrophil seg.', 'Eosinophil',
           'Basophil', 'Other', 'ESR', 'CK-MB'],  'A1.')

code_1(1, ['Free T4', 'TSH', 'T3', 'T4',
           'Thyroglobulin Ag', 'Thyroglobulin Ab', 'Anti-TPO (TPO-Ab)', 'TSH-R-Ab', 'TBPE 정성',
           '25-OH Vitamin D3(RIA)', '25-OH-Vit.D','25-OH Vitamin D3 (RIA)',
           'Testosterone', 'Free testosterone', 'Vitamin B12(Cyanocobalamine)',
           'Insulin', 'C-peptide',
           'FSH', 'LH', 'E2(Estradiol)', 'Estrogen,total', 'DHEA-s',
           'Prolactin(PRL)', 'β-hCG', 'Growth Hormone', 'IGF-1(Somatomedin-C)',
           'Cortisol'], 'H1.')

code_1(1, ['ABO혈액형(자동화)', 'Rh type(자동화)'],      'C1.')

code_1(1, ['AFP', 'CEA', 'CA19-9', 'CA125',
           'PSA (total PSA)', 'Cyfra 21-1'],     'B1.')

code_1(1, ['RF 정성', 'RF 정량', 'CRP 정성', 'CRP 정량', 'RPR정밀', 'HIV Ab(AIDS Ab)',
           'HAV-Ab IgG','HAV-Ab IgM',
           'HBsAg', 'HBsAg(일반)', 'HBsAb', 'HCV Ab(일반)', 'HCV Ab', 'HCV RNA 정량',
           'HBV DNA 정량(Realtime PCR)', 'HBV DNA IU/mL', 'HBV DNA copies/mL', 'HBe Ag', 'HBe Ab',],  'D1.')

code_1(1, ['Urine 10종', 'Specific gravity', 'PH', 'Leukocytes', 'Nitrite', 'Proteins', 'Glucose',
           'Ketones', 'Urobilinogen', 'Bilirubin', 'Blood',
           '요침사검사', 'WBC', 'RBC', 'Epithelial cells', 'Bacteria', 'Crystals', 'Casts', 'Others',
           'Creatinine(Urine)', 'Microalbumin', 'A/C ratio'],    'K1.')

code_1(1, ['Stool Hb 정량', 'Egg count', 'Microscopy(S)', 'Stool parasite','Ova.',
           '분변충란검사(집란법)', 'Stool WBC count'],   'S1.')

code_1(1, ['Level B', 'Gross Findings', 'Microscopic Findings', 'Diagnosis', 'Comment',
           '일반 부인과 검사(Pap stain, GY)','1. SPECIMEN ADEQUACY', '2. GENERAL CATEGORIZATION',
           '3. INTERPRETATION AND RESULT', '4. COMMENTS',
           'Level B(공단)', 'Level C(Blocks≤9)'],     'T1.')

code_1(1, ['Rubella IgG', 'Measles IgG', 'Mumps IgG',
           'Eosinophil count',
           'PT', 'PT(Sec)', 'PT(%)', 'PT(INR)', 'APTT',
           'Magnesium(Mg)', 'Magnesium(Mg)',
           'P.B Cell Morphology (PBS)', 'WBC_Number', 'WBC_Maturation',
           'WBC_Neutro - Segment', 'WBC_Neutro - Toxic gran', 'WBC_Distribution', 'RBC_Size',
           'RBC_Chromicity', 'RBC_Anisocytosis', 'RBC_Poikilocytosis', 'RBC_Elliptocyte', 'RBC_Burr cell',
           'RBC_Target cell', 'RBC_Spherocyte', 'RBC_Schistocyte', 'RBC_Dimorphism', 'RBC_Rouleaux formation',
           'RBC_Polychromasia', 'RBC_Basophilic stippling', 'RBC_Howell - Jolly Body', 'RBC_Nucleated RBC',
           'Platelet_Number', 'Platelet_Size', 'Platelet_Clumping', 'Reticulocyte count',
           'CBC 8종',
           'M.tuberculosis PCR', 'QuantiFERON-TB', 'AFB stain(Direct)', 'AFB culture',
           'VZV IgG(Varicella Zoster)',
           'Slide 대여'], 'Z_Etc.')

code_1(41, ['MAST 93', 'Total IgE(총IgE)', 'D.pteronyssinus(유럽집먼지진드기)', 'D.farinae(미국집먼지진드기)',
'Acarus siro(수중다리가루진드기)', 'Storage mite(저장진드기)', 'Cat(고양이)', 'Horse(말)', 'Dog(개)', 'Guinea pig(기니피그)',
'Sheep(양)', 'Rabbit(토끼)', 'Hamster(햄스터)', 'Egg white(계란흰자)', 'Milk(우유)', 'Codfish(대구)', 'Wheat flour(밀가루)',
'Barley meal(보리)', 'Maize(옥수수)', 'Rice(쌀)', 'Sesame(참깨)', 'Buck-wheat(메밀)', 'Peanut(땅콩)', 'Soy bean(콩)', 'Crab(게)',
'Shrimp(새우)', 'Tomato(토마토)', 'Pork(돼지고기)', 'Beef(소고기)', 'Potato(감자)', 'Mussel(홍합)', 'Tuna(참치)', 'Salmon(연어)',
'Yeast, bakers(효모)', 'Garlic(마늘)', 'Onion(양파)', 'Apple(사과)', 'Cheddar cheese(치즈)', 'Chicken(닭고기)', 'Kiwi(키위)',
'Celery(샐러리)', 'Mango(망고)', 'Banana(바나나)', 'Cacao(카카오)', 'Peach(복숭아)', 'Mackerel(고등어)', 'Clam(조개)',
'Mushroom(버섯)', 'Cucumber(오이)', 'Walnut(호두)', 'Squid(오징어)', 'Chestnut(밤)', 'Anchovy(멸치)',
'Citrus mix(감귤류(레몬,오렌지))', 'Sweet vernal grass(향기풀)', 'Rye pollens(호밀풀)',
'Bermuda grass(우산잔디)', 'Orchard grass(오리새)', 'Timothy grass(큰조아재비)', 'Reed(갈대)', 'Redtop, bent grass(외겨이삭)',
'House dust(집먼지)', 'Honey bee(꿀벌)', 'Yellow jacket (wasp)(말벌)', 'Cockroach(바퀴벌레)', 'Penicillium notatum(곰팡이류)',
'Cladosporium herbarum(곰팡이류)', 'Aspergillus fumigatus(곰팡이류)', 'Candida albicans(곰팡이류)', 'Alternaria alternata(곰팡이류)',
'Alder(오리나무)', 'Birch(자작나무)', 'Hazel(개암나무)', 'Oak white(참나무)', 'Sycamore mix(플라타너스)', 'Sallow willow(수양버들)',
'Poplar mix(포플라)', 'Ash mix(물푸레)', 'Pine(소나무)', 'Japanese cedar(삼나무)', 'Acacia(아카시아)', 'Hinoki cypress(편백나무)',
'Ragweed, short(돼지풀)', 'Mugwort(쑥)', 'Oxeye daisy(블란서 국화)', 'Dandelion(민들레)', 'English plantain(창질경이)',
'Russian thistle(명아주과풀)', 'Goldenrod(미역취 국화)', 'Pigweed(털비름)', 'Japanese hop(환삼덩굴)', 'Latex(라텍스)',
'Pupa, silk cocoon(번데기)', 'CCD(Bromelain)(파인애플 브로멜라인)'], 'Z_MAST.')


df2 = df1.sort_values(by='dt_code_1')
# df2.drop_duplicates(subset= '수진자명' , keep= False, inplace=True)
df2_len = (len(df2))
name_list = df2['수진자명'].values.tolist()

df7 = df2
u = 0
for i in name_list:
    u=u+1
    print(u)
    m1 = df7[(df7['검체명'].isin(['Tissue', 'Tissue B', 'Tissue C-9', 'Vaginal smear.',
                             'Urine (Random)', 'Stool']) == False) & df7['수진자명'].isin([i])]
    m2 = df7[df7['검체명'].isin(['Tissue', 'Tissue B', 'Tissue C-9']) & df7['수진자명'].isin([i])]
    m3 = df7[df7['검체명'].isin(['Vaginal smear.']) & df7['수진자명'].isin([i])]
    m4 = df7[df7['검체명'].isin(['Urine (Random)']) & df7['수진자명'].isin([i])]
    m5 = df7[df7['검체명'].isin(['Stool']) & df7['수진자명'].isin([i])]

    # writer = pd.ExcelWriter('C:/GDSRC_2019/Data/Ianuarius/extraction/2019/' +
    #                         'D_2019_' + i + '_.xlsx', engine='xlsxwriter')

    writer = pd.ExcelWriter('C:/GDSRC/Data/Ianuarius/extraction/2020/' +
                            'D_2020_' + i + '_.xlsx', engine='xlsxwriter')

    m1.to_excel(writer, sheet_name='Laboratory', index=False)
    m2.to_excel(writer, sheet_name='Tissue', index=False)
    m3.to_excel(writer, sheet_name='PAP', index=False)
    m4.to_excel(writer, sheet_name='Urine', index=False)
    m5.to_excel(writer, sheet_name='Stool', index=False)
    print(i)
    writer.save()
# df2.to_excel('./GDS_Lab_ 2001_2006.xlsx')
