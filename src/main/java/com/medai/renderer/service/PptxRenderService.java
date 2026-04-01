package com.medai.renderer.service;

import com.medai.renderer.model.ConfidenceScore;
import com.medai.renderer.model.RenderRequest;
import com.medai.renderer.model.SlideData;
import com.medai.renderer.template.ThemeConfig;
import com.medai.renderer.util.PptxUtils;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.stereotype.Service;
import java.awt.Dimension;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

@Service
public class PptxRenderService {
    private final ChartService chartService;
    public PptxRenderService(ChartService chartService) { this.chartService = chartService; }

    public byte[] render(RenderRequest request) throws Exception {
        XMLSlideShow pptx = new XMLSlideShow();
        if (request.isWidescreen()) pptx.setPageSize(new Dimension(960, 540));
        for (SlideData sd : request.getSlides()) {
            XSLFSlide slide = pptx.createSlide();
            String layout = sd.getLayout() != null ? sd.getLayout() : "CONTENT_FULL";
            PptxUtils.setBackground(slide, ThemeConfig.bgColorFor(layout, request.getTheme()));
            switch (layout) {
                case "TITLE"           -> buildTitle(slide, pptx, sd, request);
                case "TOC"             -> buildToc(slide, sd, request);
                case "DIVIDER"         -> buildDivider(slide, sd);
                case "CONTENT_FULL"    -> buildContent(slide, sd);
                case "CONTENT_TWO_COL" -> buildTwoCol(slide, sd);
                case "CONTENT_CARDS"   -> buildCards(slide, sd);
                case "TABLE"           -> buildTable(slide, sd);
                case "CHART_KM"        -> buildKmChart(slide, pptx, sd);
                case "SWOT"            -> buildSwot(slide, pptx, sd);
                case "TIMELINE"        -> buildTimeline(slide, sd);
                case "KPI_DASHBOARD"   -> buildKpi(slide, sd);
                case "REFERENCES"      -> buildRefs(slide, sd);
                case "CONFIDENCE"      -> buildConfidence(slide, sd, request);
                default                -> buildContent(slide, sd);
            }
        }
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        pptx.write(baos); pptx.close();
        return baos.toByteArray();
    }

    private void buildTitle(XSLFSlide sl, XMLSlideShow pptx, SlideData d, RenderRequest req) {
        PptxUtils.addRect(sl, 0, 0, 13.33, 0.05, ThemeConfig.HEX_ACCENT);
        PptxUtils.addRect(sl, 0, 7.45, 13.33, 0.05, ThemeConfig.HEX_TEAL);
        PptxUtils.addRect(sl, 0, 0.05, 0.08, 7.40, ThemeConfig.HEX_TEAL);
        PptxUtils.addRect(sl, 0.5, 0.8, 12.33, 3.2, ThemeConfig.HEX_SURFACE);
        PptxUtils.addText(sl, d.contentStr("title"), 0.8, 1.0, 11.0, 1.4, 40.0, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE, true, TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(sl, d.contentStr("subtitle"), 0.8, 2.5, 11.0, 0.7, 18.0, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED, false, TextParagraph.TextAlign.LEFT);
        @SuppressWarnings("unchecked") List<Object> badges = (List<Object>) d.getContent().getOrDefault("badges", List.of());
        double bx = 0.8;
        for (Object badge : badges) { String t = badge != null ? badge.toString() : ""; if (t.isEmpty()) continue; double tw = Math.max(1.8, t.length()*0.13+0.4); PptxUtils.addRect(sl, bx, 3.35, tw, 0.35, ThemeConfig.HEX_NAVY); PptxUtils.addText(sl, t, bx, 3.35, tw, 0.35, 9.0, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEAL, true, TextParagraph.TextAlign.CENTER); bx += tw + 0.15; }
        List<Map<String, Object>> kpis = d.contentList("kpis"); double kw=3.6, kx=0.8;
        for (int i = 0; i < Math.min(kpis.size(), 3); i++) { Map<String,Object> k = kpis.get(i); String v=k.getOrDefault("value","").toString(), l=k.getOrDefault("label","").toString(), c=resolveColor(k.getOrDefault("color","accent").toString()); PptxUtils.addRect(sl, kx, 4.6, kw, 1.5, ThemeConfig.HEX_SURFACE); PptxUtils.addRect(sl, kx, 4.6, kw, 0.05, c); PptxUtils.addText(sl, v, kx, 4.7, kw, 0.85, 32.0, ThemeConfig.FONT_TITLE, c, true, TextParagraph.TextAlign.CENTER); PptxUtils.addText(sl, l, kx, 5.5, kw, 0.45, 11.0, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED, false, TextParagraph.TextAlign.CENTER); kx += kw + 0.25; }
        PptxUtils.addText(sl, "Generated: "+req.meta("generatedAt")+" | MedAI Suite DeepResearch", 0.8, 6.7, 11.0, 0.3, 8.0, ThemeConfig.FONT_BODY, ThemeConfig.HEX_DIM, false, TextParagraph.TextAlign.LEFT);
    }

    private void buildToc(XSLFSlide sl, SlideData d, RenderRequest req) {
        PptxUtils.addHeader(sl, "Contents", req.meta("drug"));
        List<Map<String,Object>> items = d.contentList("items"); int cb=(items.size()+1)/2; double sy=1.2;
        for (int i = 0; i < items.size(); i++) { String t=items.get(i).getOrDefault("title","").toString(); boolean lc=i<cb; double x=lc?0.7:6.8, y=sy+(lc?i:i-cb)*0.50; PptxUtils.addRect(sl,x,y+0.03,0.38,0.35,ThemeConfig.HEX_ACCENT); PptxUtils.addText(sl,String.valueOf(i+1),x,y+0.03,0.38,0.35,10.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_WHITE,true,TextParagraph.TextAlign.CENTER); PptxUtils.addText(sl,t,x+0.50,y,5.2,0.42,12.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.LEFT); }
        PptxUtils.addFooter(sl, null);
    }

    private void buildDivider(XSLFSlide sl, SlideData d) {
        PptxUtils.addRect(sl,0,0,13.33,0.05,ThemeConfig.HEX_ACCENT); PptxUtils.addRect(sl,0,7.45,13.33,0.05,ThemeConfig.HEX_TEAL);
        PptxUtils.addText(sl,String.format("%02d",d.getSectionIndex()),0.6,1.0,4.0,3.0,96.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_SURFACE,true,TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(sl,d.contentStr("title"),0.8,3.2,10.0,1.3,36.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_WHITE,true,TextParagraph.TextAlign.LEFT);
        PptxUtils.addRect(sl,0.8,4.5,2.5,0.05,ThemeConfig.HEX_TEAL);
        String sub=d.contentStr("subtitle"); if(!sub.isEmpty()) PptxUtils.addText(sl,sub,0.8,4.7,10.0,0.6,16.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_MUTED,false,TextParagraph.TextAlign.LEFT);
    }

    private void buildContent(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl, d.contentStr("title"), d.getSection());
        String body=d.contentStr("body");
        if(!body.isEmpty()) {
            // Detect if body contains bullet-like lines
            String[] lines = body.split("\n");
            boolean hasBullets = false;
            for (String line : lines) {
                if (line.trim().startsWith("•") || line.trim().startsWith("▸") || line.trim().matches("^\\d+\\.\\s.*")) {
                    hasBullets = true; break;
                }
            }
            if (hasBullets && lines.length > 2) {
                // Render as multi-paragraph with proper spacing
                PptxUtils.addText(sl, body, 0.6, 1.15, 12.1, 5.2, 11.5, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT, false, TextParagraph.TextAlign.LEFT);
            } else {
                PptxUtils.addText(sl, body, 0.6, 1.15, 12.1, 5.2, 12.5, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT, false, TextParagraph.TextAlign.LEFT);
            }
        }
        List<Map<String,Object>> items=d.contentList("items");
        if(!items.isEmpty()&&body.isEmpty()) {
            String[] t=items.stream().map(m->m.getOrDefault("text",m.toString()).toString()).toArray(String[]::new);
            PptxUtils.addBulletText(sl,t,0.6,1.15,12.1,5.2,12.0,ThemeConfig.HEX_TEXT);
        }
        addRefFooter(sl,d); PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildTwoCol(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        PptxUtils.addText(sl,d.contentStr("leftTitle"),0.6,1.15,5.8,0.45,14.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_TEAL,true,TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(sl,d.contentStr("leftBody"),0.6,1.65,5.8,4.6,11.5,ThemeConfig.FONT_BODY,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.LEFT);
        PptxUtils.addRect(sl,6.6,1.15,0.03,5.2,ThemeConfig.HEX_ACCENT);
        PptxUtils.addText(sl,d.contentStr("rightTitle"),6.9,1.15,5.8,0.45,14.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_ACCENT,true,TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(sl,d.contentStr("rightBody"),6.9,1.65,5.8,4.6,11.5,ThemeConfig.FONT_BODY,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.LEFT);
        addRefFooter(sl,d); PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildCards(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        List<Map<String,Object>> cards=d.contentList("cards");
        int count=Math.min(cards.size(),6);
        if(count==0){buildContent(sl,d);return;}
        int cols = count <= 2 ? 2 : count <= 4 ? 2 : 3;
        int rows = (count + cols - 1) / cols;
        double gap = 0.18;
        double cW = (12.1 - (cols - 1) * gap) / cols;
        double availH = 5.5;  // Content area height
        double cH = (availH - (rows - 1) * gap) / rows;

        for (int i = 0; i < count; i++) {
            Map<String,Object> c = cards.get(i);
            int col = i % cols, row = i / cols;
            double cx = 0.6 + col * (cW + gap);
            double cy = 1.15 + row * (cH + gap);
            String ac = ThemeConfig.ACCENT_CYCLE[i % ThemeConfig.ACCENT_CYCLE.length];

            // Card background
            PptxUtils.addRect(sl, cx, cy, cW, cH, ThemeConfig.HEX_SURFACE);
            // Top accent bar
            PptxUtils.addRect(sl, cx, cy, cW, 0.06, ac);
            // Title — larger, accent-colored
            PptxUtils.addText(sl, c.getOrDefault("title","").toString(),
                cx + 0.18, cy + 0.14, cW - 0.36, 0.45,
                14.0, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE, true,
                TextParagraph.TextAlign.LEFT);
            // Body — slightly larger, better wrapping
            PptxUtils.addText(sl, c.getOrDefault("body","").toString(),
                cx + 0.18, cy + 0.62, cW - 0.36, cH - 0.80,
                11.0, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT, false,
                TextParagraph.TextAlign.LEFT);
        }
        addRefFooter(sl,d); PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildTable(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        Map<String,Object> td=d.contentMap("table"); if(td.isEmpty()){buildContent(sl,d);return;}
        @SuppressWarnings("unchecked") List<String> h=((List<Object>)td.getOrDefault("headers",List.of())).stream().map(Object::toString).toList();
        @SuppressWarnings("unchecked") List<List<String>> r=((List<List<Object>>)td.getOrDefault("rows",List.of())).stream().map(row->row.stream().map(Object::toString).toList()).toList();
        if(h.isEmpty()){buildContent(sl,d);return;} String[] hA=h.toArray(new String[0]); int mx=r.size(); String[][] rA=new String[mx][]; for(int i=0;i<mx;i++)rA[i]=r.get(i).toArray(new String[0]);
        PptxUtils.addTable(sl,hA,rA,0.5,1.15,12.3,0.38);
        addRefFooter(sl,d); PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildKmChart(XSLFSlide sl, XMLSlideShow pptx, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        Map<String,Object> cd=d.contentMap("chartData"); if(!cd.isEmpty()){try{byte[] png=chartService.generateKaplanMeier(cd);PptxUtils.addImage(sl,pptx,png,0.5,1.10,12.3,5.3);}catch(Exception e){PptxUtils.addText(sl,"Chart error: "+e.getMessage(),0.6,3.0,12.1,1.0,12.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_ROSE,false,TextParagraph.TextAlign.CENTER);}}
        addRefFooter(sl,d); PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildSwot(XSLFSlide sl, XMLSlideShow pptx, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        Map<String,Object> sw=d.contentMap("swotData"); if(!sw.isEmpty()){try{byte[] png=chartService.generateSwotMatrix(sw);PptxUtils.addImage(sl,pptx,png,0.8,1.10,11.7,5.3);}catch(Exception e){PptxUtils.addText(sl,"SWOT error: "+e.getMessage(),0.6,3.0,12.1,1.0,12.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_ROSE,false,TextParagraph.TextAlign.CENTER);}}
        PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildTimeline(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        List<Map<String,Object>> ev=d.contentList("events"); if(ev.isEmpty())return; int n=Math.min(ev.size(),7); double barY=3.2, sp=11.7/Math.max(n,1);
        PptxUtils.addRect(sl,0.8,barY,11.7,0.06,ThemeConfig.HEX_ACCENT);
        for(int i=0;i<n;i++){Map<String,Object> e=ev.get(i);double cx=0.8+i*sp+sp/2;String ac=ThemeConfig.ACCENT_CYCLE[i%ThemeConfig.ACCENT_CYCLE.length];PptxUtils.addRect(sl,cx-0.10,barY-0.10,0.26,0.26,ac);PptxUtils.addText(sl,e.getOrDefault("date","").toString(),cx-0.7,barY-0.85,1.6,0.55,11.0,ThemeConfig.FONT_TITLE,ac,true,TextParagraph.TextAlign.CENTER);PptxUtils.addText(sl,e.getOrDefault("title","").toString(),cx-0.7,barY+0.30,1.6,0.45,10.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_WHITE,true,TextParagraph.TextAlign.CENTER);PptxUtils.addText(sl,e.getOrDefault("detail","").toString(),cx-0.7,barY+0.75,1.6,1.2,8.5,ThemeConfig.FONT_BODY,ThemeConfig.HEX_MUTED,false,TextParagraph.TextAlign.CENTER);}
        PptxUtils.addFooter(sl,null);
    }

    private void buildKpi(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl,d.contentStr("title"),d.getSection());
        List<Map<String,Object>> kpis=d.contentList("kpis"); int cols=Math.min(kpis.size(),4); if(cols==0){buildContent(sl,d);return;} double kW=(12.1-(cols-1)*0.20)/cols, kH=2.2;
        for(int i=0;i<Math.min(kpis.size(),8);i++){Map<String,Object> k=kpis.get(i);int col=i%cols,row=i/cols;double kx=0.6+col*(kW+0.20),ky=1.25+row*(kH+0.20);String ac=ThemeConfig.ACCENT_CYCLE[i%ThemeConfig.ACCENT_CYCLE.length];PptxUtils.addRect(sl,kx,ky,kW,kH,ThemeConfig.HEX_SURFACE);PptxUtils.addRect(sl,kx,ky,kW,0.05,ac);PptxUtils.addText(sl,k.getOrDefault("value","—").toString(),kx,ky+0.20,kW,1.2,36.0,ThemeConfig.FONT_TITLE,ac,true,TextParagraph.TextAlign.CENTER);PptxUtils.addText(sl,k.getOrDefault("label","").toString(),kx+0.1,ky+1.45,kW-0.2,0.6,9.5,ThemeConfig.FONT_BODY,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.CENTER);}
        addRefFooter(sl,d); PptxUtils.addFooter(sl,confBadge(d));
    }

    private void buildRefs(XSLFSlide sl, SlideData d) {
        PptxUtils.addHeader(sl,"References",d.getSection());
        List<Map<String,Object>> refs=d.contentList("references"); double y=1.15;
        for(int i=0;i<refs.size()&&y<6.6;i++){Map<String,Object> r=refs.get(i);String t="["+(i+1)+"] "+r.getOrDefault("text","").toString();int tier=r.containsKey("tier")?((Number)r.get("tier")).intValue():4;String tc=tier==1?ThemeConfig.HEX_TEAL:tier==2?ThemeConfig.HEX_ACCENT:tier==3?ThemeConfig.HEX_GOLD:ThemeConfig.HEX_DIM;PptxUtils.addRect(sl,0.6,y+0.06,0.10,0.10,tc);PptxUtils.addText(sl,t,0.8,y,12.0,0.28,8.5,ThemeConfig.FONT_MONO,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.LEFT);y+=0.32;}
        PptxUtils.addFooter(sl,null);
    }

    private void buildConfidence(XSLFSlide sl, SlideData d, RenderRequest req) {
        PptxUtils.addRect(sl,0,0,13.33,0.05,ThemeConfig.HEX_TEAL); PptxUtils.addRect(sl,0,7.45,13.33,0.05,ThemeConfig.HEX_TEAL);
        ConfidenceScore cs=req.getConfidenceScore(); if(cs==null)cs=new ConfidenceScore();
        PptxUtils.addText(sl,"Confidence Score: "+cs.getOverall()+"% ("+cs.getGrade()+")",0.8,0.4,10.0,0.9,32.0,ThemeConfig.FONT_TITLE,ThemeConfig.HEX_TEAL,true,TextParagraph.TextAlign.LEFT);
        String[][] sc={{"Source Verification",cs.getSourceVerification()+"%","Weight: 35%",ThemeConfig.HEX_TEAL},{"Traceability",cs.getTraceability()+"%","Weight: 30%",ThemeConfig.HEX_ACCENT},{"Source Quality",cs.getSourceQuality()+"%","Weight: 20%",ThemeConfig.HEX_GOLD},{"Cross-Reference",cs.getCrossReference()+"%","Weight: 15%",ThemeConfig.HEX_ROSE}};
        double bx=0.6; for(String[] s:sc){PptxUtils.addRect(sl,bx,1.65,3.0,2.0,ThemeConfig.HEX_SURFACE);PptxUtils.addRect(sl,bx,1.65,3.0,0.05,s[3]);PptxUtils.addText(sl,s[1],bx,1.85,3.0,0.95,30.0,ThemeConfig.FONT_TITLE,s[3],true,TextParagraph.TextAlign.CENTER);PptxUtils.addText(sl,s[0],bx,2.80,3.0,0.40,10.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.CENTER);PptxUtils.addText(sl,s[2],bx,3.20,3.0,0.30,8.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_DIM,false,TextParagraph.TextAlign.CENTER);bx+=3.15;}
        String gt=req.meta("generationTimeSeconds"); if(!gt.isEmpty()){try{int s=Integer.parseInt(gt);PptxUtils.addText(sl,"Generated in "+(s/60)+"m "+(s%60)+"s  |  Industry: 2-4 weeks  |  Savings: ~99.9%",0.6,4.1,12.1,0.5,13.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_TEXT,false,TextParagraph.TextAlign.LEFT);}catch(NumberFormatException ignored){}}
        PptxUtils.addText(sl,"Estimated Cost Savings: EUR 3,200-6,400 per MAP  |  80-160h x EUR 40/h avg.",0.6,4.65,12.1,0.45,12.0,ThemeConfig.FONT_BODY,ThemeConfig.HEX_GOLD,false,TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(sl,"Formula: SV x 0.35 + TR x 0.30 + SQ x 0.20 + CR x 0.15 | Tiers: T1=PubMed (100%), T2=Conference (85%), T3=Guideline (70%), T4=Other (40%) | EU AI Act: LIMITED-RISK",0.6,5.5,12.1,0.8,8.0,ThemeConfig.FONT_MONO,ThemeConfig.HEX_DIM,false,TextParagraph.TextAlign.LEFT);
    }

    private String resolveColor(String key){return switch(key.toLowerCase()){case "accent"->ThemeConfig.HEX_ACCENT;case "teal"->ThemeConfig.HEX_TEAL;case "gold"->ThemeConfig.HEX_GOLD;case "rose"->ThemeConfig.HEX_ROSE;case "orange"->ThemeConfig.HEX_ORANGE;default->key.length()==6?key:ThemeConfig.HEX_ACCENT;};}

    private void addRefFooter(XSLFSlide sl, SlideData d) {
        List<Map<String,Object>> refs=d.contentList("references"); if(refs.isEmpty())return;
        StringBuilder sb=new StringBuilder(); for(int i=0;i<Math.min(refs.size(),5);i++){if(i>0)sb.append("  |  ");sb.append("[").append(i+1).append("] ");String t=refs.get(i).getOrDefault("text","").toString();sb.append(t.length()>80?t.substring(0,77)+"...":t);}
        PptxUtils.addText(sl,sb.toString(),0.5,6.60,12.3,0.30,7.0,ThemeConfig.FONT_MONO,ThemeConfig.HEX_DIM,false,TextParagraph.TextAlign.LEFT);
    }

    private String confBadge(SlideData d){List<Map<String,Object>> r=d.contentList("references");if(r.isEmpty())return null;long t1=r.stream().filter(x->x.containsKey("tier")&&((Number)x.get("tier")).intValue()==1).count();return "Sources: "+r.size()+" ("+t1+" Gold) | MedAI Suite";}
}
