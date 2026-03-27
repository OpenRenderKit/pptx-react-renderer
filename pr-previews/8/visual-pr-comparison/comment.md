<!-- pptx-react-renderer-visual-pr-comment -->
## Visual PR Diff

Compared `15` visual cases against the latest successful `main` visual artifact and found `8` changed cases.

- [Workflow run](https://github.com/OpenRenderKit/pptx-react-renderer/actions/runs/23630479419)
- [PR vs main report](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/index.html)
- [Office reference vs current report](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-regression/index.html)

| Case | Max Diff | Mean Diff | Top Changed Slides |
| --- | ---: | ---: | --- |
| complex-sanitized | 0.4025 | 0.1506 | 4 (0.4025), 5 (0.1224), 8 (0.0505), 1 (0.0268) |
| slop-sample | 0.3953 | 0.0564 | 4 (0.3953), 5 (0.1168), 10 (0.0613), 8 (0.0429), 11 (0.0250) |
| table-layout | 0.1811 | 0.1811 | 1 (0.1811) |
| theme-luminance | 0.1750 | 0.1750 | 1 (0.1750) |
| theme-colors | 0.1125 | 0.1125 | 1 (0.1125) |
| diagram-fallback | 0.0620 | 0.0310 | 10 (0.0620), 2 (0.0000) |
| group-transform | 0.0591 | 0.0591 | 1 (0.0591) |
| nested-groups | 0.0150 | 0.0150 | 1 (0.0150) |
| bullets-numbering | 0.0000 | 0.0000 | 1 (0.0000) |
| image-cropping | 0.0000 | 0.0000 | 1 (0.0000) |

### Preview Slides

**complex-sanitized** slide 4 ([open case](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/complex-sanitized/comparison.json))

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/complex-sanitized/slide-04.baseline.png" width="260" alt="Main baseline complex-sanitized slide 4"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/complex-sanitized/slide-04.current.png" width="260" alt="PR current complex-sanitized slide 4"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/complex-sanitized/slide-04.diff.png" width="260" alt="Diff complex-sanitized slide 4"></td></tr>
</table>

**slop-sample** slide 4 ([open case](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/slop-sample/comparison.json))

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/slop-sample/slide-04.baseline.png" width="260" alt="Main baseline slop-sample slide 4"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/slop-sample/slide-04.current.png" width="260" alt="PR current slop-sample slide 4"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/slop-sample/slide-04.diff.png" width="260" alt="Diff slop-sample slide 4"></td></tr>
</table>

**table-layout** slide 1 ([open case](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/table-layout/comparison.json))

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/table-layout/slide-01.baseline.png" width="260" alt="Main baseline table-layout slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/table-layout/slide-01.current.png" width="260" alt="PR current table-layout slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/8/visual-pr-comparison/table-layout/slide-01.diff.png" width="260" alt="Diff table-layout slide 1"></td></tr>
</table>

Artifacts include HTML side-by-side reports under `visual-regression-artifacts` and `visual-pr-comparison`.