import zipfile
import os
import shutil

output_dir = "C:/Users/User/createslide/assets"
os.makedirs(output_dir, exist_ok=True)

# Map: (pptx_file, internal_media_name) -> output_name
extractions = [
    # Demo ① 天母場地概覽 (研發討論 Slide 1 = media2.mp4)
    ("C:/Users/User/createslide/研發討論會議_訓練負荷主題1217.pptx", "ppt/media/media2.mp4", "demo1_tianmu_overview.mp4"),
    # Demo ② 球道追蹤實測 (115年度 Slide 5 = media2.mp4)
    ("C:/Users/User/createslide/115年度棒球智慧訓練規劃V3.pptx", "ppt/media/media2.mp4", "demo2_ball_tracking.mp4"),
    # Demo ③ 出手點辨識 (研發討論 Slide 2 = media1.mp4)
    ("C:/Users/User/createslide/研發討論會議_訓練負荷主題1217.pptx", "ppt/media/media1.mp4", "demo3_release_point.mp4"),
    # Demo ④ 數據查詢操作 (科專成果 Slide 8 = media1.mp4)
    ("C:/Users/User/createslide/科專成果技術導入天母棒球場說明簡報_工研院中分院20251013_Tie.pptx", "ppt/media/media1.mp4", "demo4_data_query.mp4"),
    # Demo ⑤ 洲際場系統 1
    ("C:/Users/User/createslide/洲際棒球場.pptx", "ppt/media/media1.mp4", "demo5_zhongji_1.mp4"),
    # Demo ⑥ 洲際場系統 2
    ("C:/Users/User/createslide/洲際棒球場.pptx", "ppt/media/media2.mp4", "demo6_zhongji_2.mp4"),
    # Demo ⑦ 洲際場系統 3
    ("C:/Users/User/createslide/洲際棒球場.pptx", "ppt/media/media3.mp4", "demo7_zhongji_3.mp4"),
    # Demo ⑧ 完整訓練情境 (115年度 Slide 1 = media1.mp4)
    ("C:/Users/User/createslide/115年度棒球智慧訓練規劃V3.pptx", "ppt/media/media1.mp4", "demo8_training_scenario.mp4"),
    # Demo ⑨ 完整動作捕捉 (115年度 Slide 9 = media3.mp4)
    ("C:/Users/User/createslide/115年度棒球智慧訓練規劃V3.pptx", "ppt/media/media3.mp4", "demo9_motion_capture.mp4"),
]

for pptx_path, internal_path, output_name in extractions:
    output_path = os.path.join(output_dir, output_name)
    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            if internal_path in z.namelist():
                with z.open(internal_path) as src, open(output_path, 'wb') as dst:
                    shutil.copyfileobj(src, dst)
                size_mb = os.path.getsize(output_path) / 1024 / 1024
                print(f"OK: {output_name} ({size_mb:.1f} MB)")
            else:
                available = [f for f in z.namelist() if 'media' in f and f.endswith('.mp4')]
                print(f"NOT FOUND: {internal_path} in {os.path.basename(pptx_path)}")
                print(f"  Available mp4s: {available}")
    except Exception as e:
        print(f"ERROR: {output_name}: {e}")

print("\nDone. Files in assets:")
for f in sorted(os.listdir(output_dir)):
    size = os.path.getsize(os.path.join(output_dir, f)) / 1024 / 1024
    print(f"  {f} ({size:.1f} MB)")
