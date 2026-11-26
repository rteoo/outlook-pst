# Project artwork

`outlook-pst-icon.svg` is the icon's source: a mail envelope with a search lens
at its corner, on a neutral ink tile. `outlook-pst-icon.png` is its 1024-pixel
render, used by the root README and the plugin manifests. Its transparent outer
margin works on light and dark backgrounds. No Microsoft or Outlook mark is used.

## Design

The icon follows the personal-plugin icon family, built on the Teô Design
System 3.2.0 mark (`T.`, a letter with an amber full stop): a flat `#0A0A0A` ink
tile, a 24-unit outline pictogram in `#F5F5F5` with rounded caps and joins, and a
single amber `#FE9A00` accent, here the search lens. No gradients, sheen, or
glow. Redrawn on 2026-11-26, replacing the navy-gradient version.

## Regenerating the PNG

Render the SVG at 1024 by 1024 with a transparent background, for example with
headless Microsoft Edge:

```text
msedge --headless=new --disable-gpu --hide-scrollbars --default-background-color=00000000 --window-size=1024,1024 --screenshot=outlook-pst-icon.png outlook-pst-icon.svg
```

The plugin build copies the PNG byte for byte, so rebuild the plugin after
changing it.
