const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  AlignmentType, LevelFormat, BorderStyle, TabStopType,
} = require("docx");

const LINK = "0000EE";
const FONT = "Cambria";

const numbering = {
  config: [{
    reference: "bullets",
    levels: [{
      level: 0,
      format: LevelFormat.BULLET,
      text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: {
        paragraph: {
          indent: { left: 540, hanging: 270 },
        },
      },
    }],
  }],
};

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 2 } },
    children: [new TextRun({ text, bold: true, size: 26, font: FONT })],
  });
}

function titleLine(text) {
  return new Paragraph({
    spacing: { before: 140, after: 0 },
    children: [new TextRun({ text, bold: true, size: 22, font: FONT })],
  });
}

function companyDateLine(company, location, dateRange) {
  return new Paragraph({
    spacing: { before: 20, after: 40 },
    tabStops: [{ type: TabStopType.RIGHT, position: 10080 }],
    children: [
      new TextRun({ text: company, bold: true, size: 21, font: FONT }),
      new TextRun({ text: ` (${location})`, size: 21, font: FONT }),
      new TextRun({ text: "\t", size: 21, font: FONT }),
      new TextRun({ text: dateRange, italics: true, size: 21, font: FONT }),
    ],
  });
}

function bullet(parts) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 15, after: 15 },
    children: parts.map(r => {
      if (typeof r === "string") return new TextRun({ text: r, size: 21, font: FONT });
      if (r instanceof ExternalHyperlink) return r;
      return new TextRun({ ...r, size: 21, font: FONT });
    }),
  });
}

function link(text, url) {
  return new ExternalHyperlink({
    link: url,
    children: [new TextRun({ text, size: 21, font: FONT, color: LINK, underline: {} })],
  });
}

function projectTitleLine(name, dateRange) {
  return new Paragraph({
    spacing: { before: 140, after: 40 },
    tabStops: [{ type: TabStopType.RIGHT, position: 10080 }],
    children: [
      new TextRun({ text: name, bold: true, size: 22, font: FONT }),
      new TextRun({ text: "\t", size: 21, font: FONT }),
      new TextRun({ text: dateRange, italics: true, size: 21, font: FONT }),
    ],
  });
}

const doc = new Document({
  numbering,
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 1080, bottom: 720, left: 1080 },
      },
    },
    children: [

      // ═══ NAME ═══
      new Paragraph({
        spacing: { after: 40 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Aaron Collins", bold: true, size: 48, font: FONT })],
      }),

      // ═══ CONTACT ═══
      new Paragraph({
        spacing: { before: 20, after: 0 },
        alignment: AlignmentType.CENTER,
        children: [
          link("aaron777collins@gmail.com", "mailto:aaron777collins@gmail.com"),
          new TextRun({ text: "  |  ", size: 20, font: FONT }),
          new TextRun({ text: "+1 (517) 515-0233", size: 20, font: FONT }),
          new TextRun({ text: "  |  ", size: 20, font: FONT }),
          link("LinkedIn", "https://www.linkedin.com/in/aaron-collins-ab70511b7/"),
          new TextRun({ text: "  |  ", size: 20, font: FONT }),
          link("aaroncollins.info", "https://www.aaroncollins.info/"),
        ],
      }),

      // ═══ EXPERIENCE ═══
      sectionHeading("Experience"),

      // -- Senior SDE --
      titleLine("Senior Software Development Engineer"),
      companyDateLine("Criteo", "Ann Arbor, MI", "Apr 2026 - Present"),
      bullet([
        "Retail Media / Onsite Ads / Activation SME: Driving major re-architecture of ad approval systems to improve configuration flexibility. Own Ads \"Activation\" (",
        { text: "ad validation and delivery push", bold: true },
        ") pipelines, backend, and UI. Collaborate cross-team for ad delivery workflows (C#/.NET 10, Angular)."
      ]),
      bullet([
        "Project SME and core contributor to the Ad Diagnostic tool MVP (identifying why configured ads do not serve), designed to significantly reduce troubleshooting time and help advertisers get campaigns live faster."
      ]),
      bullet([
        "Selected as AI Champion by senior leadership: driving AI adoption, onboarding engineers to AI tooling and workflows, and tracking cross-org AI initiatives. Measured via dedicated OKR."
      ]),

      // -- SDE II --
      titleLine("Software Development Engineer II"),
      companyDateLine("Criteo", "Ann Arbor, MI", "Oct 2024 - Apr 2026"),
      bullet([
        "Retail Media / Onsite Ads / Fullstack Engineering Project SME: Led major overhaul of Ads \"Forecasting\" (",
        { text: "ad delivery demand prediction", bold: true },
        ") system and built new modules that significantly improved forecast accuracy. Owned pipeline, backend infrastructure, and business requirements implementation. Collaborated cross-team for ad approval workflows (C#/.NET 8-10, Angular)."
      ]),
      bullet([
        "Owned full project lifecycles from architecture through delivery, driving measurable business impact across cross-domain Onsite Ads initiatives."
      ]),
      bullet([
        "Built ",
        { text: "ForgeBot", bold: true },
        ", a Slack-integrated AI coding agent (Anthropic Agent SDK). Engineers interact with it directly in Slack channels to fix bugs, add features, or summarize topics, reducing context-switching overhead."
      ]),
      bullet([
        "Developed ",
        { text: "Claude Code Helper", bold: true },
        ", a CLI harness wrapping Claude Code with experimental agent teams, multi-threaded conversations with thread tracking, pre/post-instruction injection, and post-hook scripting. Accessible via private Slack DM. Boosted developer productivity."
      ]),

      // -- SDE --
      titleLine("Software Development Engineer"),
      companyDateLine("Criteo", "Ann Arbor, MI", "Oct 2023 - Oct 2024"),
      bullet([
        "Delivered high-value Onsite Display Ads features including UserNotification, Auction BuyType (auction-based bidding for display line items), and Max Scale (multi-creative support, expanded configuration). Drove strong stakeholder engagement (C#/.NET 6-8, Angular)."
      ]),
      bullet([
        "Resolved critical production bugs with potential for significant financial impact across ad serving systems."
      ]),
      bullet([
        "Authored operational documentation, Grafana dashboards, and alerting systems, transforming incident response. Go-to knowledge source for Criteo's async AWS messaging infrastructure."
      ]),

      // -- Rocket --
      titleLine("Software Developer Co-op (2x)"),
      companyDateLine("Rocket Innovation Studio", "Windsor, ON", "May - Aug 2021, Jan - May 2022"),
      bullet([
        "Developed front-end components in Angular 11/AngularJS with full unit test coverage and backend services in .NET Core 3/5/6."
      ]),

      // -- TA & Research --
      titleLine("Head Teaching Assistant & Lead Research Assistant"),
      companyDateLine("University of Windsor", "Windsor, ON", "Sep 2021 - Aug 2023"),
      bullet([
        "Head TA for the Python programming course; also served as TA for one additional course."
      ]),
      bullet([
        "Led connected/autonomous driving research using machine learning. Managed and mentored a team of research assistants."
      ]),
      bullet([
        "Published cognitive load recognition paper (Sci-Kit Learn, eye-tracking) at ACR\u201923, published by Springer. ",
        link("Cognitive Load Recognition via Eye-Tracking", "https://doi.org/10.1007/978-3-031-33743-7_7"),
      ]),
      bullet([
        "Built configurable ML pipeline for BSM classification in connected driving (Sci-Kit Learn, TensorFlow, Pandas). ",
        link("ConnectedDrivingPipelineV4", "https://www.aaroncollins.info/ConnectedDrivingPipelineV4/"),
      ]),

      // ═══ PROJECTS (page 2) ═══
      new Paragraph({
        children: [new (require("docx").PageBreak)()],
      }),
      sectionHeading("Projects"),

      projectTitleLine("PortableRalph", "Jan 2026 - Present"),
      bullet([
        "Autonomous AI development loop that works in any repo. Cross-platform (Linux, macOS, Windows), 13+ GitHub stars. ",
        link("GitHub", "https://github.com/aaron777collins/portableralph"),
        " | ",
        link("Docs", "https://www.portableralph.com/"),
      ]),

      projectTitleLine("SOPHIE", "2025 - Present"),
      bullet([
        "Personal AI life assistant accessible via Slack, Discord, and voice calls. Manages emails, preps tax documents, makes phone calls, and works on side projects autonomously using skills/MCP integrations. ",
        link("sophie", "https://www.aaroncollins.info/projects/sophie.html"),
        " | ",
        link("sophie-mcp", "https://www.aaroncollins.info/projects/sophie-mcp.html"),
        " | ",
        link("sophie-voice-bridge", "https://www.aaroncollins.info/projects/sophie-voice-bridge.html"),
      ]),

      projectTitleLine("Minecraft Datapack Language (MDL)", "2024 - Present"),
      bullet([
        "Compiler and programming language for Minecraft datapacks with if/else, loops, and variables. Published on PyPI with a VS Code extension. ",
        link("GitHub", "https://github.com/aaron777collins/MinecraftDatapackLanguage"),
        " | ",
        link("Docs", "https://www.aaroncollins.info/MinecraftDatapackLanguage/"),
      ]),

      projectTitleLine("Control Windows", "Jun 2020 - Jan 2023"),
      bullet([
        "Desktop automation tool using Python and PyAutoGUI with auto-code generation capabilities. ",
        link("GitHub", "https://github.com/aaron777collins/Control-Windows"),
        " | ",
        link("Docs", "https://www.aaroncollins.info/projects/Control-Windows.html"),
      ]),

      projectTitleLine("ZeroTier-GUI-Arch", "2023 - Present"),
      bullet([
        "ZeroTier GUI client for Arch Linux / Steam Deck. 41+ GitHub stars, popular in the Steam Deck community with active releases via Flatpak. ",
        link("GitHub", "https://github.com/aaron777collins/ZeroTier-GUI-Arch"),
        " | ",
        link("Docs", "https://www.aaroncollins.info/projects/ZeroTier-GUI-Arch.html"),
      ]),

      new Paragraph({
        spacing: { before: 80, after: 0 },
        indent: { left: 270 },
        children: [
          new TextRun({ text: "More projects at ", italics: true, size: 20, font: FONT }),
          link("aaroncollins.info/projects", "https://www.aaroncollins.info/projects/"),
          new TextRun({ text: " | ", italics: true, size: 20, font: FONT }),
          link("github.com/aaron777collins", "https://github.com/aaron777collins/"),
        ],
      }),

      // ═══ EDUCATION ═══
      sectionHeading("Education"),

      titleLine("Bachelor of Computer Science (Honours)"),
      new Paragraph({
        spacing: { before: 20, after: 40 },
        tabStops: [{ type: TabStopType.RIGHT, position: 10080 }],
        children: [
          new TextRun({ text: "University of Windsor", bold: true, size: 21, font: FONT }),
          new TextRun({ text: "\t", size: 21, font: FONT }),
          new TextRun({ text: "2023", italics: true, size: 21, font: FONT }),
        ],
      }),
      bullet(["94% Major Average (With Great Distinction)"]),
      bullet(["President\u2019s Renewable Entrance Scholarship (95% Entry Average)"]),
      bullet(["NSERC Undergraduate Student Research Award"]),

      // ═══ SKILLS ═══
      sectionHeading("Skills"),

      bullet([
        { text: "Languages: ", bold: true },
        "C#, Python, Java, TypeScript, JavaScript, C, R, SQL, Bash"
      ]),
      bullet([
        { text: "Frameworks & Web: ", bold: true },
        ".NET (Core 3-10), Angular, React, Flask, Node.js, HTML, CSS/SCSS"
      ]),
      bullet([
        { text: "AI & ML: ", bold: true },
        "Anthropic Agent SDK, Claude Code, AMP, Cursor, MCP, Ollama, TensorFlow, Sci-Kit Learn, Pandas, NumPy"
      ]),
      bullet([
        { text: "Infrastructure & DevOps: ", bold: true },
        "Docker, AWS (SQS, Lambda, ECS), Jenkins, Grafana, Kibana, NGINX, Caddy, Authelia (SSO), CI/CD"
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Resume_Aaron_Collins.docx", buffer);
  console.log("Done");
});
