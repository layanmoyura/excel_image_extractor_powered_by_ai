"""Generates a realistic multi-sheet sample Excel file containing the user's
exact 62-row jewelry catalog data with embedded product pictures.
"""

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage, ImageDraw, ImageFont
import io

USER_CSV_DATA = """Serial No ,id,,Design Number,Image,Category,Metal,Colour,Gross weight,Net weight ,Total Net weight,,Stone Category,Stone type,Stone count,Stone weight,Total Stone weight,Stone Category,Stone type,Stone count,Stone weight,Total Stone weight,Cost in US $,Quantity,Total
1,377393,,18259W1BS8WV,,Earring,18K,White,1.1148,1.03,2.06,,Diamond,white dia,24,0.062,0.124,Precious,Blue Sapphire,2,0.3,0.6,295,2,590
2,373226,,70002QBS8WP,,Ring,18K,White,2.437,2.28,4.56,EU 54,Diamond,white dia,12,0.119,0.238,Precious,Blue Sapphire,1,0.547,1.094,615,2,1230
3,373248,,77172QBS8WV,,Ring,18K,White,1.7296,1.62,1.62,EU 50,Diamond,white dia,37,0.116,0.116,Precious,Blue Sapphire,1,0.432,0.432,404,1,404
4,373307,,77172W1BS8WV,,Earring,18K,White,1.3398,1.22,1.22,,Diamond,white dia,66,0.149,0.149,Precious,Blue Sapphire,2,0.45,0.45,367,1,367
5,373308,,77171W1BS8WV,,Earring,18K,White,1.4358,1.29,2.58,,Diamond,white dia,64,0.142,0.284,Precious,Blue Sapphire,2,0.445,0.89,373,2,746
6,373353,,18250WBS8WV,,Earring,18K,White,1.1146,1.02,2.04,,Diamond,white dia,14,0.039,0.078,Precious,Blue Sapphire,2,0.395,0.79,319,2,638
7,373354,,18254QBS8WV,,Ring,18K,White,1.2338,1.16,2.32,EU 50,Diamond,white dia,14,0.033,0.066,Precious,Blue Sapphire,1,0.303,0.606,301,2,602
8,373889,,77171UBS8WV,,Necklace,18K,White,2.1364,2.01,2.01,18",Diamond,white dia,32,0.11,0.11,Precious,Blue Sapphire,1,0.522,0.522,494,1,494
9,373244,,18248W2BS8WV,,Earring,18K,White,1.1708,1.08,2.16,,Diamond,white dia,22,0.05,0.1,Precious,Blue Sapphire,2,0.354,0.708,321,2,642
10,375432,,18259QBS8WV,,Ring,18K,White,1.144,1.08,4.32,EU50X2,Diamond,white dia,12,0.029,0.116,Precious,Blue Sapphire,1,0.204,0.816,252,4,1008
11,373367,,18248SBS8WV,,Pendant,18K,White,0.4354,0.37,0.74,,Diamond,white dia,13,0.033,0.066,Precious,Blue Sapphire,1,0.261,0.522,179,2,358
12,375544,,60498W1BS8WP,,Earring,18K,White,1.4056,1.27,1.27,,Diamond,white dia,32,0.179,0.179,Precious,Blue Sapphire,2,0.499,0.499,514,1,514
13,373168,,18250S1BS8WV,,Pendant,18K,White,0.4502,0.39,0.78,,Diamond,white dia,7,0.019,0.038,Precious,Blue Sapphire,1,0.263,0.526,174,2,348
14,373185,,18273S1BS8WV,,Pendant,18K,White,0.5652,0.51,0.51,,Diamond,white dia,20,0.053,0.053,Precious,Blue Sapphire,1,0.223,0.223,198,1,198
15,373228,,76225QBS8WV,,Ring,18K,White,1.6468,1.54,1.54,EU 52,Diamond,white dia,20,0.05,0.05,Precious,Blue Sapphire,1,0.484,0.484,434,1,434
16,375214,,18248W2BS8WV,,Earring,18K,White,1.1874,1.1,1.1,,Diamond,white dia,22,0.052,0.052,Precious,Blue Sapphire,2,0.385,0.385,336,1,336
17,373318,,18757QBS8WP,,Ring,18K,White,1.5338,1.49,1.49,EU 52,Diamond,white dia,12,0.061,0.061,Precious,Blue Sapphire,1,0.158,0.158,298,1,298
18,373342,,18259SBS8WV,,Pendant,18K,White,0.4554,0.41,0.41,,Diamond,white dia,12,0.03,0.03,Precious,Blue Sapphire,1,0.197,0.197,130,1,130
19,375217,,18757SBS8WP,,Pendant,18K,White,0.4044,0.36,0.36,,Diamond,white dia,12,0.062,0.062,Precious,Blue Sapphire,1,0.16,0.16,137,1,137
20,373345,,18757WBS8WP,,Earring,18K,White,0.8998,0.85,0.85,,Diamond,white dia,24,0.089,0.089,Precious,Blue Sapphire,2,0.16,0.16,257,1,257
21,373361,,71544SBS8WV,,Pendant,18K,White,0.6976,0.57,1.14,,Diamond,white dia,29,0.075,0.15,Precious,Blue Sapphire,1,0.488,0.976,318,2,636
22,375546,,77171QBS8WV,,Ring,18K,White,1.6836,1.55,1.55,EU 50,Diamond,white dia,32,0.096,0.096,Precious,Blue Sapphire,1,0.572,0.572,488,1,488
23,375493,,18248SBS8WV,,Pendant,18K,White,0.4486,0.39,0.39,,Diamond,white dia,13,0.032,0.032,Precious,Blue Sapphire,1,0.261,0.261,182,1,182
24,374907,,77171UBS8YV,,Necklace,18K,Yellow,2.2116,2.1,4.2,18",Diamond,white dia,32,0.095,0.19,Precious,Blue Sapphire,1,0.368,0.736,469,2,938
25,375565,,60498SBS8WP,,Pendant,18K,White,0.7428,0.63,0.63,,Diamond,white dia,22,0.144,0.144,Precious,Blue Sapphire,1,0.42,0.42,357,1,357
26,374918,,71544W1BS8WV,,Earring,18K,White,1.5362,1.37,2.74,,Diamond,white dia,40,0.07,0.14,Precious,Blue Sapphire,2,0.691,1.382,501,2,1002
27,375208,,12151SBS8WV,,Pendant,18K,White,0.5906,0.51,0.51,,Diamond,white dia,16,0.053,0.053,Precious,Blue Sapphire,1,0.35,0.35,193,1,193
28,373164,,70002QBS8WP,,Ring,18K,White,2.3102,2.12,6.36,EU50X2,Diamond,white dia,12,0.131,0.393,Precious,Blue Sapphire,1,0.558,1.674,602,3,1806
29,373165,,70002SBS8WP,,Pendant,18K,White,0.8224,0.63,3.15,,Diamond,white dia,12,0.08,0.4,Precious,Blue Sapphire,1,0.562,2.81,361,5,1805
30,375213,,15697SBS8WV,,Pendant,18K,White,0.4022,0.35,0.7,,Diamond,white dia,12,0.055,0.11,Precious,Blue Sapphire,1,0.151,0.302,155,2,310
31,375433,,18757QBS8WP,,Ring,18K,White,1.533,1.48,2.96,EU50X2,Diamond,white dia,12,0.055,0.11,Precious,Blue Sapphire,1,0.155,0.31,292,2,584
32,375434,,18757WBS8WP,,Earring,18K,White,0.9234,0.85,1.7,,Diamond,white dia,24,0.089,0.178,Precious,Blue Sapphire,2,0.189,0.378,229,2,458
33,375491,,13905W1BS8WV,,Earring,18K,White,1.2942,1.15,5.75,,Diamond,white dia,40,0.102,0.51,Precious,Blue Sapphire,2,0.211,1.055,279,5,1395
34,375497,,18254QBS8WV,,Ring,18K,White,1.267,1.18,4.72,EU50X2,Diamond,white dia,14,0.031,0.124,Precious,Blue Sapphire,1,0.311,1.244,256,4,1024
35,375530,,72564UBS8RV,,Necklace,18K,Rose,2.2406,2.11,2.11,18",Diamond,white dia,8,0.019,0.019,Precious,Blue Sapphire,5,0.634,0.634,471,1,471
36,375547,,77171UBS8WV,,Necklace,18K,White,2.1004,2,2,18",Diamond,white dia,32,0.102,0.102,Precious,Blue Sapphire,1,0.4,0.4,529,1,529
37,375567,,18254SBS8WV,,Pendant,18K,White,0.4614,0.37,1.85,,Diamond,white dia,14,0.031,0.155,Precious,Blue Sapphire,1,0.302,1.51,146,5,730
38,375572,,18259UBS8WV,,Necklace,18K,White,1.924,1.88,1.88,18",Diamond,white dia,12,0.031,0.031,Precious,Blue Sapphire,1,0.189,0.189,377,1,377
39,374299,,12151W1BS8WV,,Earring,18K,White,1.098,0.95,4.75,,Diamond,white dia,28,0.064,0.32,Precious,Blue Sapphire,2,0.42,2.1,271,5,1355
40,374300,,77171W1BS8YV,,Earring,18K,Yellow,1.4814,1.28,5.12,,Diamond,white dia,64,0.144,0.576,Precious,Blue Sapphire,2,0.431,1.724,437,4,1748
41,374894,,18254WBS8WV,,Earring,18K,White,1.095,0.97,3.88,,Diamond,white dia,28,0.05,0.2,Precious,Blue Sapphire,2,0.425,1.7,269,4,1076
42,374901,,77171QBS8YV,,Ring,18K,Yellow,1.874,1.7,6.8,EU50X2,Diamond,white dia,32,0.103,0.412,Precious,Blue Sapphire,1,0.458,1.832,417,4,1668
43,377354,,72564WBS8RV,,Earring,18K,Rose,1.4578,1.31,1.31,,Diamond,white dia,8,0.026,0.026,Precious,Blue Sapphire,6,0.713,0.713,364,1,364
44,377375,,77171UBS8YV,,Necklace,18K,Yellow,2.2692,2.07,10.35,18",Diamond,white dia,32,0.103,0.515,Precious,Blue Sapphire,1,0.481,2.405,491,5,2455
45,377379,,71544W1BS8RV,,Earring,18K,Rose,1.438,1.29,1.29,,Diamond,white dia,40,0.071,0.071,Precious,Blue Sapphire,2,0.669,0.669,431,1,431
46,377396,,70002WBS8WP,,Earring,18K,White,1.6374,1.34,8.04,,Diamond,white dia,24,0.159,0.954,Precious,Blue Sapphire,2,0.533,3.198,506,6,3036
47,377401,,71544SBS8RV,,Pendant,18K,Rose,0.7262,0.54,3.24,,Diamond,white dia,29,0.081,0.486,Precious,Blue Sapphire,1,0.445,2.67,293,6,1758
48,,,BP49008R-20WD8RP,,Ring,18K,Rose,2.6226,2.43,7.29,EU50X1,Diamond,white dia,21,0.321,0.963,,,,,,691,3,2073
49,,,BP49008R-20WD8YP,,Ring,18K,Yellow,2.5812,2.45,4.9,EU50X1,Diamond,white dia,21,0.328,0.656,,,,,,701,2,1402
50,,,BP48733W-10EM8YP,,Earring,18K,Yellow,2.7242,2.42,9.68,,Diamond,white dia,78,0.333,1.332,Precious,Emerald,2,0.189,0.756,732,4,2928
51,,,BP48813QBS8WP,,Ring,18K,White,4.0706,3.41,23.87,EU50X2,Diamond,white dia,18,0.437,3.059,Precious,Blue Sapphire,1,0.244,1.708,938,7,6566
52,,,26024ULB8WX,,Necklace,18K,White,1.064,0.96,0.96,18",,,,,,Semi Precious,London Blue,1,0.52,0.52,183,1,183
53,,,26024WLB8WX,,Earring,18K,White,1.022,0.84,0.84,,,,,,,Semi Precious,London Blue,2,0.91,0.91,164,1,164
54,,,DI45882NVR8RP,,Necklace,18K,Rose,1.145,1.09,1.09,18",Diamond,white dia,4,0.03,0.03,Precious,Ruby,1,0.245,0.245,259,1,259
55,,,DI45882NVE8YP,,Necklace,18K,Yellow,1.168,1.13,1.13,18",Diamond,white dia,4,0.027,0.027,Precious,Emerald,1,0.163,0.163,251,1,251
56,,,BP47629WWD8YP,,Earring,18K,Yellow,1.0172,0.99,1.98,,Diamond,white dia,2,0.068,0.136,,,,,,232,2,464
57,,,77171QBS8WV,,Ring,18K,White,1.7028,1.55,3.1,EU 50,Diamond,white dia,32,0.096,0.192,Precious,Blue Sapphire,1,0.572,1.144,413,2,826
58,377380,,77171W1BS8WV,,Earring,18K,White,1.3322,1.2,1.2,,Diamond,white dia,64,0.144,0.288,Precious,Blue Sapphire,2,0.373,0.746,347,1,347
59,,,FI53648BWD8YP,,Bracelet,18K,Yellow,4.248,4.2,4.2,,Diamond,white dia,58,0.24,0.24,,,,,,816,1,816
60,,,BP48742QBS8WP,,Ring,18K,White,3.7352,3.63,3.63,EU 52,Diamond,white dia,28,,0.526,Precious,Blue Sapphire,1,,0.258,951,1,951
61,,,BP48742UBS8WP,,Necklace,18K,White,5.0452,4.98,4.98,,Diamond,white dia,17,,0.326,Precious,Blue Sapphire,1,,0.277,1097,1,1097
62,,,BP48742WBS8WP,,Earring,18K,White,2.875,2.75,2.75,,Diamond,white dia,24,,0.625,Precious,Blue Sapphire,2,,0.496,1016,1,1016
"""

def generate_sample_file(output_filename="Jewelry_Catalog_Sample.xlsx"):
    print(f"Generating realistic workbook: {output_filename}...")
    wb = openpyxl.Workbook()
    
    # Sheet 1: Main Jewelry Catalog (User's Exact Format)
    ws1 = wb.active
    ws1.title = "Jewelry Orders"
    
    # Rows 1-2 blank
    # Row 3 Category headers
    ws1.cell(row=3, column=13, value="Diamond")
    ws1.cell(row=3, column=18, value="Precious")
    
    csv_lines = USER_CSV_DATA.strip().split("\n")
    headers = csv_lines[0].split(",")
    
    # Headers on Row 4
    for col_idx, h in enumerate(headers, 1):
        ws1.cell(row=4, column=col_idx, value=h.strip())
        
    # Populate data rows 5 to 66
    data_lines = csv_lines[1:]
    for idx, line in enumerate(data_lines):
        r_num = 5 + idx
        parts = line.split(",")
        design_no = ""
        category = ""
        
        for col_idx, val in enumerate(parts, 1):
            val_clean = val.strip()
            if col_idx == 4:  # Design Number column
                design_no = val_clean
            elif col_idx == 6: # Category
                category = val_clean
                
            if val_clean:
                # Try float
                try:
                    num_val = float(val_clean)
                    ws1.cell(row=r_num, column=col_idx, value=num_val)
                except ValueError:
                    ws1.cell(row=r_num, column=col_idx, value=val_clean)
                    
        # Generate an appealing jewelry dummy image
        img = PILImage.new("RGB", (90, 90), color=(240, 243, 246))
        draw = ImageDraw.Draw(img)
        # Draw nice jewel ring/pendant motif
        draw.rectangle([5, 5, 84, 84], outline=(200, 205, 215), width=2)
        accent_color = (60, 120, 216) if "Sapphire" in line else (220, 60, 60) if "Ruby" in line else (46, 160, 90) if "Emerald" in line else (200, 160, 50)
        draw.ellipse([25, 18, 65, 58], outline=accent_color, width=3)
        draw.rectangle([40, 12, 50, 22], fill=accent_color)
        
        # Draw design number label on thumbnail
        short_label = design_no[:8] if design_no else f"ITEM_{idx+1}"
        draw.text((8, 66), short_label, fill=(40, 40, 50))
        
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        
        xl_img = XLImage(buf)
        # Place in Column E (col 5) at row r_num
        ws1.add_image(xl_img, f"E{r_num}")
        
    # Sheet 2: Rings Special Catalog (Alternative column layout to showcase flexibility!)
    ws2 = wb.create_sheet(title="Custom Rings")
    ws2.cell(row=2, column=2, value="Style SKU")
    ws2.cell(row=2, column=3, value="Metal")
    ws2.cell(row=2, column=5, value="Jewelry Photo")
    
    custom_designs = [
        ("70002QBS8WP-A", "18K White"),
        ("77172QBS8WV-B", "18K Rose"),
        ("BP49008R-20WD8RP-C", "18K Yellow")
    ]
    
    for c_idx, (sku, metal) in enumerate(custom_designs):
        r = 3 + c_idx
        ws2.cell(row=r, column=2, value=sku)
        ws2.cell(row=r, column=3, value=metal)
        
        img = PILImage.new("RGB", (90, 90), color=(250, 245, 235))
        draw = ImageDraw.Draw(img)
        draw.ellipse([20, 20, 70, 70], outline=(212, 175, 55), width=4)
        draw.text((10, 72), sku[:10], fill=(60, 50, 20))
        
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        xl_img = XLImage(buf)
        ws2.add_image(xl_img, f"E{r}")
        
    wb.save(output_filename)
    print(f"Successfully generated '{output_filename}' with {len(data_lines)} images on Sheet 1 and {len(custom_designs)} on Sheet 2.")

if __name__ == "__main__":
    generate_sample_file()
