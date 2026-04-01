<!-- pptx-react-renderer-visual-pr-comment -->
## Visual PR Diff

Compared `15` visual cases against the latest successful `main` visual artifact and found `2` changed cases.

- [Workflow run](https://github.com/OpenRenderKit/pptx-react-renderer/actions/runs/23839477430)
- [PR vs main report](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html)
- [Office reference vs current report](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-regression/index.html)

### Renderer Change Ranking

| Case | Max Diff | Mean Diff | Top Changed Slides |
| --- | ---: | ---: | --- |
| [nested-groups](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-nested-groups) | 0.0481 | 0.0481 | 1 (0.0481) |
| [group-transform](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-group-transform) | 0.0411 | 0.0411 | 1 (0.0411) |
| [bullets-numbering](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-bullets-numbering) | 0.0000 | 0.0000 | 1 (0.0000) |
| [complex-sanitized](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-complex-sanitized) | 0.0000 | 0.0000 | 1 (0.0000), 4 (0.0000), 5 (0.0000), 8 (0.0000) |
| [diagram-fallback](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-diagram-fallback) | 0.0000 | 0.0000 | 2 (0.0000), 10 (0.0000) |
| [image-cropping](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-image-cropping) | 0.0000 | 0.0000 | 1 (0.0000) |
| [image-placement](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-image-placement) | 0.0000 | 0.0000 | 1 (0.0000) |
| [paragraph-spacing](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-paragraph-spacing) | 0.0000 | 0.0000 | 1 (0.0000) |
| [slop-sample](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-slop-sample) | 0.0000 | 0.0000 | 1 (0.0000), 2 (0.0000), 3 (0.0000), 4 (0.0000), 5 (0.0000) |
| [table-layout](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-table-layout) | 0.0000 | 0.0000 | 1 (0.0000) |

### Most Regressed Vs Office Reference

| Case | Mean Delta | Max Delta | Worst Slide Delta |
| --- | ---: | ---: | --- |
| [group-transform](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-regression/index.html#case-group-transform) | +0.0001 | +0.0001 | 1 (+0.0001) |

### Most Improved Vs Office Reference

| Case | Mean Delta | Max Delta | Best Slide Delta |
| --- | ---: | ---: | --- |
| [nested-groups](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-regression/index.html#case-nested-groups) | -0.0000 | -0.0000 | 1 (-0.0000) |
| [text-insets-anchor](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-regression/index.html#case-text-insets-anchor) | -0.0000 | -0.0000 | 1 (-0.0000) |

### Preview Slides

**[nested-groups](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-nested-groups)** slide 1

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/nested-groups/slide-01.baseline.png" width="260" alt="Main baseline nested-groups slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/nested-groups/slide-01.current.png" width="260" alt="PR current nested-groups slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/nested-groups/slide-01.diff.png" width="260" alt="Diff nested-groups slide 1"></td></tr>
</table>

**[group-transform](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/index.html#case-group-transform)** slide 1

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/group-transform/slide-01.baseline.png" width="260" alt="Main baseline group-transform slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/group-transform/slide-01.current.png" width="260" alt="PR current group-transform slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/16/visual-pr-comparison/group-transform/slide-01.diff.png" width="260" alt="Diff group-transform slide 1"></td></tr>
</table>

Artifacts include HTML side-by-side reports under `visual-regression-artifacts` and `visual-pr-comparison`.