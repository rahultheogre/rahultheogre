const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, HeadingLevel, BorderStyle, PageBreak
} = require('docx');
const fs = require('fs');

const FONT = "Arial";
const LINE_RULE = { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 };

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 180, after: 60 },
    border: { bottom: LINE_RULE },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 24, font: FONT })]
  });
}

function jobHeader(title, company, duration) {
  return new Paragraph({
    spacing: { before: 140, after: 40 },
    children: [
      new TextRun({ text: title, bold: true, size: 22, font: FONT }),
      new TextRun({ text: "  |  " + company, size: 22, font: FONT }),
      new TextRun({ text: "  |  " + duration, size: 20, italics: true, font: FONT, color: "444444" })
    ]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 30, after: 30 },
    children: [new TextRun({ text, size: 20, font: FONT })]
  });
}

function plain(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, font: FONT, ...opts })]
  });
}

function skillRow(label, value) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text: label + ": ", bold: true, size: 20, font: FONT }),
      new TextRun({ text: value, size: 20, font: FONT })
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 440, hanging: 280 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 900, right: 1080, bottom: 900, left: 1080 }
      }
    },
    children: [

      // ── NAME ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "RAHUL SHARMA", bold: true, size: 36, font: FONT })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Lead AI Engineer & Architect  |  Agentic AI & LLM Systems  |  Healthtech Founder", size: 22, font: FONT })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 160 },
        children: [
          new TextRun({ text: "rahultheogre@gmail.com  |  +91-9560825115  |  ", size: 20, font: FONT }),
          new TextRun({ text: "linkedin.com/in/rahulsharma51", size: 20, font: FONT }),
          new TextRun({ text: "  |  github.com/rahultheogre  |  Delhi-NCR (Remote Only)", size: 20, font: FONT })
        ]
      }),

      // ── SUMMARY ──
      sectionHeader("Professional Summary"),
      plain("Lead AI Engineer and Architect with 8+ years of experience building production-grade AI/ML systems across healthcare, retail, insurance, and HR-tech. Expert in Agentic AI, multi-agent orchestration (LangGraph, LangChain), LLM engineering, RAG pipelines, LLMOps, and cloud deployment on Azure and AWS. Delivered end-to-end ML systems from architecture to production across enterprise client engagements. Led teams of 8–10 engineers; experienced in stakeholder communication and client-facing delivery. Founded and built a healthtech AI platform end-to-end as sole architect and product owner — demonstrating full-stack technical leadership and independent decision-making. Strong interdisciplinary foundation: IIT Roorkee, IIIT Delhi (AI/ML), JNU (Philosophy), Hansraj College (English Literature)."),

      // ── SKILLS ──
      sectionHeader("Core Competencies"),
      skillRow("Agentic AI", "LangGraph, LangChain, Multi-Agent Systems, ReAct, Supervisor Pattern, Tool-Calling Agents, MCP (Model Context Protocol)"),
      skillRow("LLM & RAG", "Azure OpenAI, Anthropic Claude, GPT-4o, LLaMA 2, RAG (Dense + BM25 + RRF), BGE-M3, Stella400M, FlashRank, Qdrant, FAISS, Pinecone, Weaviate, Nomic AI, LambdaMART, HNSW"),
      skillRow("LLMOps & MLOps", "Langfuse, MLflow, RAGAS, DeepEval, TruLens, Azure ML, Azure AI Foundry, Databricks, Docker, Kubernetes, CI/CD, Ray, DSPy"),
      skillRow("ML & NLP", "XGBoost, CatBoost, Isolation Forest, BERT, BERTopic, SHAP, Simulated Annealing, Agglomerative Clustering, DBSCAN, NetworkX"),
      skillRow("Cloud & Infra", "Azure (Container Apps, Blob, Postgres, Key Vault), AWS (Bedrock, ECS, Lambda, DocumentDB), GitHub Actions"),
      skillRow("Tech Stack", "Python, FastAPI, Node.js, Fastify, TypeScript, React Native, Expo, Neo4j, PostgreSQL, MongoDB, Redis, WebSocket, WebRTC"),
      skillRow("AI Safety & Compliance", "SHAP explainability, PII anonymization, HIPAA-style privacy patterns, India DPDP Act 2023, RBAC, audit logging, guardrails, prompt injection defense"),

      // ── EXPERIENCE ──
      sectionHeader("Professional Experience"),

      // Freelance
      jobHeader("Lead AI Engineer & Architect", "Independent AI Consultant", "Jul 2025 – Present | Remote"),

      plain("Project 1: Agentic AI Hiring Platform", { bold: true, size: 20 }),
      bullet("Architected end-to-end multi-agent AI Hiring Platform on Azure Container Apps using LangGraph — three specialized autonomous agents: Resume Screener, Interview Analyser, and Candidate Ranker with explainable LLM-as-judge outputs"),
      bullet("Implemented RBAC across six roles, ProctorAgent and ScriptAnalyser NLP module for interview transcript analysis; MySQL schema (18 tables), FastAPI backend, React frontend, Azure AD SSO"),
      bullet("Full LLMOps observability stack: Langfuse + MLflow; deployed via GitHub Actions CI/CD with Azure Container Apps"),

      plain("Project 2: ReportSense AI — Tableau BI Report Discovery (Client: Mindbrain)", { bold: true, size: 20 }),
      bullet("Architected intelligent RAG pipeline for Tableau report discovery — 5-stage pipeline: LLM query expansion → BGE-M3 dense retrieval + BM25 lexical search → RRF fusion (dense 0.65 / BM25 0.35, k=60) → FlashRank reranking (threshold 0.65) → LLM JSON spec generation"),
      bullet("Designed modular FastAPI service with parser factory for Tableau/PowerBI; centralized prompt management; resolved token overflow via round-robin chunk distribution with 80K token hard limit"),
      bullet("Reduced LLM inference costs by ~60% through candidate pre-filtering and relevance gating before generation stage; led architecture across 3-person team"),

      plain("Project 3: AI Clinical Documentation Platform — Solo Founder & Architect (Healthtech)", { bold: true, size: 20 }),
      bullet("Founded and architected end-to-end AI-powered clinical scribe platform — sole decision-maker across product strategy, engineering, cloud infrastructure, ML/LLM pipeline, compliance, and go-to-market; clinical co-creator: Dr. Sahil Bagai, HOD Nephrology, Fortis Manesar"),
      bullet("Real-time audio capture → OpenAI Whisper STT → cost-routed LLM pipeline (Anthropic Claude for complex clinical reasoning + GPT-4o-mini for routine tasks) → structured SOAP note generation in under 60 seconds at ~Rs.1/consultation"),
      bullet("Implemented per-doctor personalization loop: correction-driven prompt enrichment adapts note style per clinician over time; specialty-aware configurable SOAP templates"),
      bullet("Built patient engagement module grounded in doctor consultation history; cross-platform (web + Android + iOS) from single React Native/Expo + TypeScript codebase; React + TypeScript web UI deployed on Azure Container Apps"),
      bullet("Designed healthcare-grade privacy architecture: PII anonymization before external LLM calls, row-level security, signed-URL audio access, full data-deletion workflow; aligned with India DPDP Act 2023"),

      plain("Project 4: AI-Powered Planogram Banding System (Retail)", { bold: true, size: 20 }),
      bullet("Architected AI Banding Engine on AKS — multi-modal embeddings (Azure OpenAI + numerical features, 80/20 weighted hybrid adjacency matrix); auto-selected clustering (Agglomerative + DBSCAN via silhouette scoring); CatBoost-based new product placement predictor with nearest-neighbour inference"),
      bullet("Processed 1,000+ products per batch in real time; reduced planogram creation time by 40% and improved placement accuracy by 25%"),
      bullet("Deployed via CI/CD on Azure DevOps with blue-green deployments, Elastic APM observability, Azure Key Vault secrets management"),

      // SRM
      jobHeader("Senior Data Scientist", "SRM Technologies", "Jul 2024 – May 2025 | Remote"),
      bullet("Built production Semantic Recommender Engine for conference networking — 3-stage pipeline: HNSW vector search (Stella400M embeddings, 768-dim, weighted Q&A 50% / Experience 30% / Education 20%) → LambdaMART Learning-to-Rank reranking → multi-feature filtering; end-to-end latency under 200ms at 1,000+ concurrent users"),
      bullet("Implemented MongoDB/DocumentDB Change Streams across 6 collection types for real-time embedding regeneration on profile updates; exponential backoff retry logic with circuit breakers"),
      bullet("Integrated AWS Bedrock for LLM inference layer; deployed on AWS ECS with auto-scaling, CloudWatch monitoring, zero-downtime rollbacks"),
      bullet("Reduced irrelevant recommendations by 75%; 85% user satisfaction on top-3 results; 40% increase in meaningful connections; 99.9% uptime"),

      // SellCX
      jobHeader("Generative AI Lead", "SellCX", "Feb 2024 – Jul 2024 | Remote"),
      bullet("Architected voice-first Agentic Sales Platform (SellBot) — LangGraph Supervisor routing 5 specialized agents (Sales, Support, Appointment, Lead Management, VCC); Deepgram STT + LiveKit WebRTC + Twilio VOIP for real-time voice pipeline with sub-second latency"),
      bullet("Built Neo4j knowledge graph across 4 data sources (conversation policies, sales rules, user profiles, conversation history); implemented Microsoft GraphRAG with Leiden community detection for hierarchical entity clustering and multi-hop reasoning"),
      bullet("Applied DSPy bootstrap few-shot optimization for prompt reliability; Pydantic-structured outputs for consistent cross-agent communication; Chain-of-Thought + ReAct prompting strategies"),
      bullet("Deployed on AWS + Kubernetes; Redis multi-layer caching, OAuth 2.0 + JWT, end-to-end voice data encryption"),

      // Status Neo
      jobHeader("Senior Consultant – AI/ML", "Status Neo Consulting", "Oct 2023 – Feb 2024 | Gurgaon"),
      bullet("Built PDF-based RAG training assistant for client Netcom — adaptive chunking (256–1024 tokens, document complexity-driven), Nomic AI embeddings, Qdrant vector store with hybrid semantic + keyword retrieval, cross-encoder reranking"),
      bullet("Implemented 3-tier caching (Redis + in-memory + file-based): 80% reduction in repeated query response time, 60% reduction in embedding compute costs"),
      bullet("Achieved 87% retrieval accuracy, 2.3s avg response time, 99.5% uptime; adopted by 90% of trainers within 3 months; reduced trainer prep time from 2 hours to 45 minutes"),
      bullet("Deployed via Docker + GitHub Actions CI/CD; React + TypeScript frontend on Azure Container Apps for trainer-facing interface"),

      // PAGE BREAK
      new Paragraph({ children: [new PageBreak()] }),

      // Onward
      jobHeader("Senior Data Scientist", "Onward Technologies", "Mar 2021 – Jul 2023 | Mumbai"),

      plain("Project 1: Catman AI Optimizer — Planogram Optimization Engine", { bold: true, size: 20 }),
      bullet("Built Catman AI Optimizer — production Simulated Annealing engine for retail shelf optimization; balanced 7 weighted objectives simultaneously: Sales, Adjacency, Utilization (overfill/underfill), Movement Cost, Inventory Gap, Empty Shelves, Constraint Warnings"),
      bullet("Designed layered architecture (Selection → Operations → Objectives → Constraints); 105+ Python files, ~15,000 LOC; extensible base-class pattern enabling new objectives without breaking existing code"),
      bullet("Implemented Ray actor-based distributed processing: 4–16 concurrent actors processing 100+ planograms in parallel; 4–5 planograms/min throughput; async WebSocket state mutations with full audit trail"),
      bullet("Achieved 25–45% improvement in planogram energy metrics; generated position/fixture-level audit reports and execution task breakdowns persisted to database"),

      plain("Project 2: NLP Insights Platform — Newspaper & Media Intelligence", { bold: true, size: 20 }),
      bullet("Built multi-faceted NLP Insights Platform for newspaper and media content analysis using BERT-based models — pipeline covered sentiment analysis, abstractive summarization, named entity recognition (NER), and topic modelling across large-scale article corpora"),
      bullet("Achieved 94% F1 on NER; deployed BERTopic for dynamic topic discovery across news categories; sentiment pipeline classified article-level and entity-level polarity for editorial and brand monitoring use cases"),
      bullet("Designed modular architecture with separate inference endpoints per NLP task; batch processing pipeline for high-volume daily news ingestion with structured output storage"),

      plain("Project 3: P&C Insurance Fraud Detection Pipeline", { bold: true, size: 20 }),
      bullet("Built end-to-end P&C insurance fraud detection system — XGBoost (supervised, weight 0.8) + Isolation Forest (unsupervised, weight 0.2) ensemble; ensemble AUC 0.87–0.96, ~10ms per-claim real-time inference"),
      bullet("Engineered 95+ features across temporal, financial, geographic, vendor, and interaction categories; added 156 graph features via NetworkX — community detection, PageRank, degree centrality for fraud ring identification"),
      bullet("Implemented SHAP explainability for individual prediction audit trails; regulatory compliance requirement for P&C insurance claims processing"),
      bullet("Flask REST API for real-time scoring; batch processing handles 10,000+ claims; under 500MB RAM for full pipeline"),

      // Twyst
      jobHeader("Data Research Analyst", "Twyst Technologies", "Aug 2016 – Apr 2020 | Bangalore"),
      bullet("Built ML models for customer segmentation, search ranking, and recommendation engines; early production exposure to data pipelines and feature engineering"),

      // ── EDUCATION ──
      sectionHeader("Education"),
      plain("PG in Computer Science and Artificial Intelligence  —  IIIT Delhi  (2021–2022)", { bold: false }),
      plain("MA Philosophy (Heidegger, Continental Philosophy)  —  Jawaharlal Nehru University (JNU)  (2016–2018)", { bold: false }),
      plain("BA (Hons) English Literature  —  Hansraj College, Delhi University  (2010–2013)", { bold: false }),
      plain("Engineering Studies  —  IIT Roorkee  (2007–2010)", { bold: false }),

      // ── CERTIFICATIONS / ADDITIONAL ──
      sectionHeader("Additional"),
      plain("Career Break (2020–2021): Independent AI/ML upskilling and preparation for postgraduate studies in AI"),
      plain("Open Source & Writing: github.com/rahultheogre  |  Medium: coffee_and_notes  |  Topics: LLMs, Agentic AI, ML Systems"),
      plain("Domain Expertise: P&C Insurance (Guidewire ecosystem), Retail Merchandising, Healthcare AI, HR-Tech"),
      plain("Interests: SLM fine-tuning (LoRA/QLoRA), LLM evaluation and oracle design, distributed AI inference (vLLM), MCP protocol"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/Rahul_Sharma_Resume.docx", buffer);
  console.log("Done");
});
