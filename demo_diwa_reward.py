import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_kit import *

# 사용: python demo_diwa_reward.py [출력.pptx]
OUT = sys.argv[1] if len(sys.argv) > 1 else "demo_diwa_reward.pptx"
d = open_deck()          # kit 폴더의 base.pptx
s = d.new_slide().header("Method (reward classifier)").page(21)

bx, by, bw, bh = BASE
VERDICT_H = 0.55
py, ph = by + MARGIN, bh - 2 * MARGIN - VERDICT_H
pw = (bw - 3 * MARGIN) / 2
L = s.panel(bx + MARGIN, py, pw, ph, title="Training (learn the judge)")
R = s.panel(bx + 2 * MARGIN + pw, py, pw, ph, title="Use (score imagined rollouts)")
lc, rc = L[0] + pw / 2, R[0] + pw / 2

# ── 왼쪽: 학습
y = [py + 1.25, py + 2.25, py + 3.3, py + 4.45, py + 5.55, py + 6.55]
s.box("demo", lc, y[0], "Expert demos")
s.box("enc", lc, y[1], "World model Encoder", icon="snow")
s.box("z", lc, y[2], "Latent z", fill="output")
s.box("pos", lc - 1.95, y[3], "Success frames")
s.box("neg", lc + 1.95, y[3], "15% random other frames")
s.box("clf", lc, y[4], r"Reward classifier $C_\psi$", fill="trainable", icon="fire")
s.box("loss", lc, y[5], r"Loss = NT-Xent (contrastive) + $\alpha$ · CE (classification)", fill="loss")

s.arrow("demo", "enc"); s.arrow("enc", "z")
zx, zy, zw, zh = s.boxes["z"]
for k, lab in (("pos", "Positive"), ("neg", "Negative")):
    kx, ky, kw, kh = s.boxes[k]; cx = kx + kw / 2
    sx = zx if cx < lc else zx + zw
    s.path([(sx, zy + zh / 2), (cx, zy + zh / 2), (cx, ky)], label=lab, label_bold=True)
cx_, cy_, cw_, ch_ = s.boxes["clf"]
for k in ("pos", "neg"):
    kx, ky, kw, kh = s.boxes[k]; cx = kx + kw / 2
    ex = cx_ if cx < lc else cx_ + cw_
    s.path([(cx, ky + kh), (cx, cy_ + ch_ / 2), (ex, cy_ + ch_ / 2)])
s.loss_link("clf", "loss")

# ── 오른쪽: 사용
s.box("zt", rc, y[0], r"Imagined latent $z_{t+1}$")
s.box("clf2", rc, y[1], r"Reward classifier $C_\psi$", fill="trainable", icon="snow")
s.box("prob", rc, y[2], r"Success probability $\in$ [0, 1]")
s.box("rew", rc, y[3] - 0.15, "Reward", sub=r"$R = C_\psi(z_{t+1})$", fill="output")
s.arrow("zt", "clf2"); s.arrow("clf2", "prob"); s.arrow("prob", "rew")
rx, ry, rw, rh = s.boxes["rew"]
s.text(rx + rw + 0.2, ry + rh / 2, "computed at every\nimagined transition", PT_BOX_SUB, True, anchor="cl")

s.verdict("The world model has **NO task reward** – this classifier creates one")
d.save(OUT)
print("saved", OUT)
