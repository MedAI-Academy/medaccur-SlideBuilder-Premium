# Template Specification for MedAI MAP Generator Premium

## How It Works

Each slide template is a standard .pptx file created in PowerPoint.
The renderer finds shapes by **name** (not by placeholder text) and fills them with data.

## Setting Shape Names in PowerPoint

1. Open PowerPoint
2. Go to **Home → Select → Selection Pane** (or Alt+F10)
3. The Selection Pane shows all shapes on the slide
4. **Double-click** a shape name to rename it
5. Use the exact names from the tables below

## Template Directory Structure

```
templates/
├── dark/                    # MedAI Dark theme
│   ├── _title.pptx          # Title slide template
│   ├── _divider.pptx        # Section divider template
│   ├── executive_summary.pptx
│   ├── prevalence_kpi.pptx
│   ├── treatment_algo.pptx
│   ├── guidelines.pptx
│   ├── pivotal_table.pptx
│   ├── swot.pptx
│   ├── imperatives.pptx
│   ├── narrative.pptx
│   ├── timeline.pptx
│   └── ...
├── light/                   # Clean Light theme
├── gray/                    # Consulting theme
├── pharma/                  # Pharma Blue theme
└── premium/                 # Black & Gold theme
```

## Required Shape Names Per Section Type

### Title Slide (_title.pptx)
| Shape Name      | Content                          | Type     |
|----------------|----------------------------------|----------|
| title_drug     | Drug name (large)                | TextBox  |
| title_subtitle | "Global/Country MAP — 2027"      | TextBox  |
| title_year     | Indication text                  | TextBox  |
| title_footer   | "MedAI Suite Premium"            | TextBox  |

### Divider Slide (_divider.pptx)
| Shape Name        | Content              | Type     |
|-------------------|----------------------|----------|
| divider_number    | Section number "01"  | TextBox  |
| divider_title     | Section title        | TextBox  |
| divider_accent_bar| Decorative bar       | Shape    |

### Executive Summary (executive_summary.pptx)
| Shape Name     | Content              | Type     |
|----------------|----------------------|----------|
| header_title   | Slide title          | TextBox  |
| card_1_title   | First card title     | TextBox  |
| card_1_body    | First card content   | TextBox  |
| card_2_title   | Second card title    | TextBox  |
| card_2_body    | Second card content  | TextBox  |
| card_3_title   | Third card title     | TextBox  |
| card_3_body    | Third card content   | TextBox  |

### Prevalence & Epidemiology (prevalence_kpi.pptx)
| Shape Name     | Content              | Type     |
|----------------|----------------------|----------|
| header_title   | Slide title          | TextBox  |
| kpi_1_value    | First KPI number     | TextBox  |
| kpi_1_label    | First KPI label      | TextBox  |
| kpi_2_value    | Second KPI number    | TextBox  |
| kpi_2_label    | Second KPI label     | TextBox  |
| kpi_3_value    | Third KPI number     | TextBox  |
| kpi_3_label    | Third KPI label      | TextBox  |

### Pivotal Studies (pivotal_table.pptx)
| Shape Name     | Content              | Type     |
|----------------|----------------------|----------|
| header_title   | Trial name + phase   | TextBox  |
| kpi_mpfs       | mPFS value           | TextBox  |
| kpi_hr         | HR value             | TextBox  |
| kpi_orr        | ORR value            | TextBox  |
| kpi_mos        | mOS value            | TextBox  |
| chart_area     | Placeholder for KM   | Shape    |

### SWOT (swot.pptx)
| Shape Name     | Content              | Type     |
|----------------|----------------------|----------|
| header_title   | Slide title          | TextBox  |
| s_title        | "Strengths"          | TextBox  |
| s_body         | Bullet points        | TextBox  |
| w_title        | "Weaknesses"         | TextBox  |
| w_body         | Bullet points        | TextBox  |
| o_title        | "Opportunities"      | TextBox  |
| o_body         | Bullet points        | TextBox  |
| t_title        | "Threats"            | TextBox  |
| t_body         | Bullet points        | TextBox  |

### Strategic Imperatives (imperatives.pptx)
| Shape Name       | Content              | Type     |
|------------------|----------------------|----------|
| header_title     | Slide title          | TextBox  |
| roof_title       | Vision statement     | TextBox  |
| pillar_1_title   | First pillar title   | TextBox  |
| pillar_1_body    | First pillar bullets | TextBox  |
| base_bar         | Foundation bar       | Shape    |

### Competitive Landscape (competitor_table.pptx)
| Shape Name     | Content              | Type     |
|----------------|----------------------|----------|
| header_title   | Slide title          | TextBox  |
| chart_area     | mPFS bar chart (PNG) | Shape    |
| table_area     | Data table           | Table    |

### Timeline (timeline.pptx)
| Shape Name     | Content              | Type     |
|----------------|----------------------|----------|
| header_title   | Slide title          | TextBox  |
| road_area      | Highway graphic      | Shape    |

## Enterprise Custom Templates

Enterprise customers can upload their own templates:

1. Download a reference template from any theme
2. Open in PowerPoint
3. Replace visual elements (colors, fonts, backgrounds, logo)
4. Keep all shape names EXACTLY as specified
5. Upload via the Premium Export screen

The `/validate-template` endpoint will check that all required shapes exist.

## Chart Insertion

Charts (KM curves, bar charts, SWOT) are rendered as PNG images by matplotlib
and inserted into the `chart_area` placeholder shape. The chart inherits the
theme's color palette automatically.

## Tips for Designers

- Use PowerPoint's built-in gradient fills, shadows, and shape effects freely
- Complex illustrations (road graphics, temple columns) should be grouped shapes
- Keep text boxes large enough — the renderer uses `word_wrap=True`
- Test with the validation endpoint before deploying
- Logo shapes should be named `logo` for enterprise branding
