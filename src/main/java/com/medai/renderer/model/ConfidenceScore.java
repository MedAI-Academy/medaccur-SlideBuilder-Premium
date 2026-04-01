package com.medai.renderer.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

@JsonIgnoreProperties(ignoreUnknown = true)
public class ConfidenceScore {

    private int overall;
    private int sourceVerification;    // SV: % of sources verified
    private int traceability;          // TR: % of claims with sources
    private int sourceQuality;         // SQ: weighted tier score
    private int crossReference;        // CR: % of claims with multi-source
    private int totalClaims;
    private int verifiedSources;
    private int totalSources;

    // Formula: overall = SV*0.35 + TR*0.30 + SQ*0.20 + CR*0.15
    public int calculateOverall() {
        return (int) Math.round(
            sourceVerification * 0.35 +
            traceability * 0.30 +
            sourceQuality * 0.20 +
            crossReference * 0.15
        );
    }

    public String getGrade() {
        if (overall >= 95) return "A+";
        if (overall >= 90) return "A";
        if (overall >= 85) return "B+";
        if (overall >= 80) return "B";
        if (overall >= 70) return "C";
        return "D";
    }

    // --- Getters & Setters ---
    public int getOverall() { return overall; }
    public void setOverall(int overall) { this.overall = overall; }
    public int getSourceVerification() { return sourceVerification; }
    public void setSourceVerification(int sv) { this.sourceVerification = sv; }
    public int getTraceability() { return traceability; }
    public void setTraceability(int tr) { this.traceability = tr; }
    public int getSourceQuality() { return sourceQuality; }
    public void setSourceQuality(int sq) { this.sourceQuality = sq; }
    public int getCrossReference() { return crossReference; }
    public void setCrossReference(int cr) { this.crossReference = cr; }
    public int getTotalClaims() { return totalClaims; }
    public void setTotalClaims(int totalClaims) { this.totalClaims = totalClaims; }
    public int getVerifiedSources() { return verifiedSources; }
    public void setVerifiedSources(int verifiedSources) { this.verifiedSources = verifiedSources; }
    public int getTotalSources() { return totalSources; }
    public void setTotalSources(int totalSources) { this.totalSources = totalSources; }
}
