# 用 Blender 的视频序列编辑器把 pptjs 渲染的帧排成介绍视频。
# blender -b --factory-startup --python make_video.py -- <帧目录> <plan.json> <输出.mp4>
# plan.json 是每张 slide 的名字（「场景:步骤」），第 i 个对应 <帧目录>/slide_<i+1>.png。
# 同一场景的步骤之间淡入（新出现的元素浮现），场景之间向左推移，开头从黑场淡入，结尾淡出到黑场。
import json
import os
import sys

import bpy

frames_dir, plan_path, out_path = sys.argv[sys.argv.index("--") + 1:]
plan = [name.split(":") for name in json.load(open(plan_path))]

FPS = 30
STEP = 24  # 中间步骤停留的帧数
FADE = 10  # 步骤之间淡入
PUSH = 20  # 场景之间推移
EDGE = 15  # 开头淡入、结尾淡出
# 每个场景最后一步停留的帧数：内容多的停久一点
LAST = {"intro": 60, "hero": 75, "formats": 75, "arch": 90, "loop": 105,
        "translate": 105, "edit": 90, "word": 90, "excel": 90, "outro": 90}

scene = bpy.context.scene
scene.render.resolution_x, scene.render.resolution_y = 1920, 1080
scene.render.resolution_percentage = 100
scene.render.fps = FPS
scene.view_settings.view_transform = "Standard"  # 颜色与 PNG 一致，不做胶片色调映射
se = scene.sequence_editor_create()

# 每一帧停留多久、和下一帧之间用什么过渡
items = []
for i, (name, _) in enumerate(plan):
    last = i + 1 == len(plan) or plan[i + 1][0] != name
    hold = LAST.get(name, 75) if last else STEP
    trans = 0 if i + 1 == len(plan) else (PUSH if last else FADE)
    items.append((hold, trans))


def key(obj, prop, frame, value):
    setattr(obj, prop, value)
    obj.keyframe_insert(prop, frame=frame)


start, prev_trans, strips = 1, 0, []
for i, (hold, trans) in enumerate(items):
    path = os.path.join(frames_dir, f"slide_{i + 1}.png")
    # 后面的帧放在更高的通道，过渡时叠在前一帧上面
    st = se.strips.new_image(f"f{i + 1}", path, i + 1, start)
    st.duration = prev_trans + hold + trans
    if i > 0:
        a, t0 = strips[-1], start
        if prev_trans == PUSH:
            key(a.transform, "offset_x", t0, 0)
            key(a.transform, "offset_x", t0 + PUSH, -1920)
            key(st.transform, "offset_x", t0, 1920)
            key(st.transform, "offset_x", t0 + PUSH, 0)
        else:
            key(st, "blend_alpha", t0, 0.0)
            key(st, "blend_alpha", t0 + FADE, 1.0)
    strips.append(st)
    start += prev_trans + hold
    prev_trans = trans

end = start - 1
# 黑场淡入、淡出（下面没有别的通道，透明处就是黑色）
first, last = strips[0], strips[-1]
key(first, "blend_alpha", 1, 0.0)
key(first, "blend_alpha", 1 + EDGE, 1.0)
key(last, "blend_alpha", end - EDGE, 1.0)
key(last, "blend_alpha", end, 0.0)

scene.frame_start, scene.frame_end = 1, end
r = scene.render
r.image_settings.media_type = "VIDEO"
r.image_settings.file_format = "FFMPEG"
r.ffmpeg.format = "MPEG4"
r.ffmpeg.codec = "H264"
r.ffmpeg.constant_rate_factor = "HIGH"
r.ffmpeg.ffmpeg_preset = "GOOD"
r.ffmpeg.gopsize = FPS * 2
r.ffmpeg.audio_codec = "NONE"
r.use_file_extension = False
r.filepath = out_path
print(f"frames: {end} ({end / FPS:.1f} s)")
bpy.ops.render.render(animation=True)
