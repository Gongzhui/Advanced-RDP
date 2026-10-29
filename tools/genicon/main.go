// Genicon draws resources/icon.png: a monitor on a blue tile.
//
//	go run ./tools/genicon
package main

import (
	"image"
	"image/color"
	"image/png"
	"log"
	"math"
	"os"
)

const size = 1024

// inRoundRect reports whether (x, y) is inside the rectangle with rounded corners.
func inRoundRect(x, y, x0, y0, x1, y1, r float64) bool {
	if x < x0 || x > x1 || y < y0 || y > y1 {
		return false
	}
	cx := math.Max(x0+r, math.Min(x, x1-r))
	cy := math.Max(y0+r, math.Min(y, y1-r))
	return (x-cx)*(x-cx)+(y-cy)*(y-cy) <= r*r
}

// cover returns the share of the pixel at (px, py) inside f, from 4×4 samples.
func cover(px, py int, f func(x, y float64) bool) float64 {
	n := 0
	for i := 0; i < 4; i++ {
		for j := 0; j < 4; j++ {
			if f(float64(px)+(float64(i)+0.5)/4, float64(py)+(float64(j)+0.5)/4) {
				n++
			}
		}
	}
	return float64(n) / 16
}

func mix(a, b color.NRGBA, t float64) color.NRGBA {
	l := func(x, y uint8) uint8 { return uint8(float64(x) + (float64(y)-float64(x))*t + 0.5) }
	return color.NRGBA{l(a.R, b.R), l(a.G, b.G), l(a.B, b.B), l(a.A, b.A)}
}

func main() {
	img := image.NewNRGBA(image.Rect(0, 0, size, size))
	top, bottom := color.NRGBA{0x3b, 0x82, 0xf6, 255}, color.NRGBA{0x1d, 0x4e, 0xd8, 255}
	white := color.NRGBA{255, 255, 255, 255}
	screen := color.NRGBA{0x0f, 0x17, 0x2a, 255}
	tile := func(x, y float64) bool { return inRoundRect(x, y, 64, 64, 960, 960, 200) }
	frame := func(x, y float64) bool { return inRoundRect(x, y, 212, 250, 812, 650, 48) }
	glass := func(x, y float64) bool { return inRoundRect(x, y, 252, 290, 772, 610, 18) }
	stand := func(x, y float64) bool {
		return inRoundRect(x, y, 462, 640, 562, 730, 0) || inRoundRect(x, y, 352, 720, 672, 774, 27)
	}
	// A chevron on the screen: the remote prompt.
	chevron := func(x, y float64) bool {
		dist := func(ax, ay, bx, by float64) float64 {
			vx, vy := bx-ax, by-ay
			t := math.Max(0, math.Min(1, ((x-ax)*vx+(y-ay)*vy)/(vx*vx+vy*vy)))
			dx, dy := x-(ax+t*vx), y-(ay+t*vy)
			return math.Hypot(dx, dy)
		}
		return dist(372, 380, 452, 450) < 22 || dist(452, 450, 372, 520) < 22 || dist(502, 520, 642, 520) < 22
	}
	for py := 0; py < size; py++ {
		bg := mix(top, bottom, float64(py)/size)
		for px := 0; px < size; px++ {
			c := color.NRGBA{}
			if a := cover(px, py, tile); a > 0 {
				c = bg
				c.A = uint8(255 * a)
				if f := cover(px, py, func(x, y float64) bool { return frame(x, y) || stand(x, y) }); f > 0 {
					c = mix(c, white, f)
				}
				if g := cover(px, py, glass); g > 0 {
					c = mix(c, screen, g)
				}
				if ch := cover(px, py, chevron); ch > 0 {
					c = mix(c, color.NRGBA{0x60, 0xa5, 0xfa, 255}, ch)
				}
			}
			img.SetNRGBA(px, py, c)
		}
	}
	f, err := os.Create("resources/icon.png")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()
	if err := png.Encode(f, img); err != nil {
		log.Fatal(err)
	}
}
