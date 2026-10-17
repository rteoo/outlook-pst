# Project artwork

`outlook-pst-icon.svg` is the icon's source: an envelope in an archive tray with
a search lens, on a rounded navy tile. `outlook-pst-icon.png` is its 1024-pixel
render, used by the root README and the plugin manifests. Its transparent outer
margin works on light and dark backgrounds. No Microsoft or Outlook mark is used.

## Design

The icon follows the personal-project icon family: a navy tile graded through
the Teô Personal Design System's navy-500, navy-600 and navy-900, a pictogram in
its paper-100 to navy-300 gradient, and a single orange-500 accent, here the
search lens. The SVG was drawn by hand on 2026-10-16, replacing the earlier
generated raster.

## Regenerating the PNG

Render the SVG at 1024 by 1024 with a transparent background, for example with
headless Microsoft Edge:

```text
msedge --headless=new --disable-gpu --hide-scrollbars --default-background-color=00000000 --window-size=1024,1024 --screenshot=outlook-pst-icon.png outlook-pst-icon.svg
```

The plugin build copies the PNG byte for byte, so rebuild the plugin after
changing it.
