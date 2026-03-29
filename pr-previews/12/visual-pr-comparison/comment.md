<!-- pptx-react-renderer-visual-pr-comment -->
## Visual PR Diff

Compared `15` visual cases against the latest successful `main` visual artifact and found `10` changed cases.

- [Workflow run](https://github.com/OpenRenderKit/pptx-react-renderer/actions/runs/23709185449)
- [PR vs main report](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html)
- [Office reference vs current report](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html)

### Renderer Change Ranking

| Case | Max Diff | Mean Diff | Top Changed Slides |
| --- | ---: | ---: | --- |
| [slop-sample](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-slop-sample) | 0.8212 | 0.1139 | 11 (0.8212), 8 (0.4539), 10 (0.0636), 9 (0.0130), 4 (0.0089) |
| [complex-sanitized](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-complex-sanitized) | 0.4743 | 0.1209 | 8 (0.4743), 4 (0.0092), 1 (0.0000), 5 (0.0000) |
| [theme-luminance](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-theme-luminance) | 0.0933 | 0.0933 | 1 (0.0933) |
| [diagram-fallback](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-diagram-fallback) | 0.0650 | 0.0347 | 10 (0.0650), 2 (0.0043) |
| [bullets-numbering](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-bullets-numbering) | 0.0106 | 0.0106 | 1 (0.0106) |
| [unsupported-chart](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-unsupported-chart) | 0.0065 | 0.0065 | 1 (0.0065) |
| [theme-colors](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-theme-colors) | 0.0059 | 0.0059 | 1 (0.0059) |
| [text-layout](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-text-layout) | 0.0054 | 0.0054 | 1 (0.0054) |
| [image-placement](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-image-placement) | 0.0052 | 0.0052 | 1 (0.0052) |
| [image-cropping](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-image-cropping) | 0.0049 | 0.0049 | 1 (0.0049) |

### Most Regressed Vs Office Reference

| Case | Mean Delta | Max Delta | Worst Slide Delta |
| --- | ---: | ---: | --- |
| [diagram-fallback](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-diagram-fallback) | +0.0031 | +0.0022 | 10 (+0.0040) |
| [nested-groups](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-nested-groups) | +0.0001 | +0.0001 | 1 (+0.0001) |

### Most Improved Vs Office Reference

| Case | Mean Delta | Max Delta | Best Slide Delta |
| --- | ---: | ---: | --- |
| [complex-sanitized](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-complex-sanitized) | -0.0952 | 0.0000 | 8 (-0.3795) |
| [slop-sample](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-slop-sample) | -0.0927 | -0.0277 | 11 (-0.7275) |
| [bullets-numbering](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-bullets-numbering) | -0.0012 | -0.0012 | 1 (-0.0012) |
| [unsupported-chart](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-unsupported-chart) | -0.0009 | -0.0009 | 1 (-0.0009) |
| [image-placement](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-regression/index.html#case-image-placement) | -0.0006 | -0.0006 | 1 (-0.0006) |

### Preview Slides

**[slop-sample](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-slop-sample)** slide 11

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/slop-sample/slide-11.baseline.png" width="260" alt="Main baseline slop-sample slide 11"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/slop-sample/slide-11.current.png" width="260" alt="PR current slop-sample slide 11"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/slop-sample/slide-11.diff.png" width="260" alt="Diff slop-sample slide 11"></td></tr>
</table>

**[complex-sanitized](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-complex-sanitized)** slide 8

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/complex-sanitized/slide-08.baseline.png" width="260" alt="Main baseline complex-sanitized slide 8"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/complex-sanitized/slide-08.current.png" width="260" alt="PR current complex-sanitized slide 8"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/complex-sanitized/slide-08.diff.png" width="260" alt="Diff complex-sanitized slide 8"></td></tr>
</table>

**[theme-luminance](https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/index.html#case-theme-luminance)** slide 1

<table>
<tr><th>Main</th><th>PR</th><th>Diff</th></tr>
<tr><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/theme-luminance/slide-01.baseline.png" width="260" alt="Main baseline theme-luminance slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/theme-luminance/slide-01.current.png" width="260" alt="PR current theme-luminance slide 1"></td><td><img src="https://openrenderkit.github.io/pptx-react-renderer/pr-previews/12/visual-pr-comparison/theme-luminance/slide-01.diff.png" width="260" alt="Diff theme-luminance slide 1"></td></tr>
</table>

Artifacts include HTML side-by-side reports under `visual-regression-artifacts` and `visual-pr-comparison`.