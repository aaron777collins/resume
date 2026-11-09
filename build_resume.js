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

      // ═══ HEADLINE ═══
      new Paragraph({
        spacing: { before: 0, after: 20 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Senior Software Engineer  |  Retail Media Ad Tech  |  AI Agents & Developer Tooling", size: 22, font: FONT })],
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
        { text: "Go-to leader for 2 of my team\u2019s 3 mission-critical ad delivery systems", bold: true },
        ", including ",
        { text: "Hermes", bold: true },
        ", which moves ",
        { text: "millions of messages", bold: true },
        " (AWS SQS) so approved ads reliably reach live auctions."
      ]),
      bullet([
        "Named ",
        { text: "AI Champion for my entire org", bold: true },
        " by senior leadership after releasing org-wide AI tools (agent harness, reusable AI skills); coach my team and help anyone who asks, ",
        { text: "driving AI adoption", bold: true },
        "."
      ]),
      bullet([
        "Co-created and now lead ",
        { text: "Servability", bold: true },
        ", our ",
        { text: "ad diagnostic (ad debug) engine", bold: true },
        " that explains why ads aren\u2019t serving, ",
        { text: "so issues get fixed faster", bold: true },
        "; expanding it to screen ads before retailer review and to ",
        { text: "decide which ads make it to delivery", bold: true },
        ", replacing a legacy library."
      ]),
      bullet([
        { text: "Designed a high-priority feature", bold: true },
        " that pauses non-paying accounts, ",
        { text: "targeting $80M+ in overdue balances across 1,400+ clients", bold: true },
        "; my design was adopted and I\u2019m helping drive delivery."
      ]),
      bullet([
        "Built a snappy ",
        { text: "Rust", bold: true },
        " app to ",
        { text: "manage and auto-purge the code workspaces AI agents spin up", bold: true },
        ", so they ",
        { text: "no longer fill developer laptops", bold: true },
        "; released to all developers."
      ]),

      // -- SDE II --
      titleLine("Software Development Engineer II"),
      companyDateLine("Criteo", "Ann Arbor, MI", "Oct 2024 - Apr 2026"),
      bullet([
        "Developed ",
        { text: "Claude Code Helper", bold: true },
        ", a Slack-run AI agent-team harness that ",
        { text: "injects user-written rules", bold: true },
        " and ships an ",
        { text: "AI brain based on Andrej Karpathy\u2019s wiki system", bold: true },
        "."
      ]),
      bullet([
        "Led the project to build a new ",
        { text: "Ads Forecasting", bold: true },
        " system (predicting ad delivery demand), ",
        { text: "owning the pipeline, backend, and business requirements", bold: true },
        " (",
        { text: "C#/.NET 8-10, Angular", bold: true },
        ")."
      ]),
      bullet([
        "Led a hackathon team to build ",
        { text: "ForgeBot", bold: true },
        ", a Slack AI coding agent (",
        { text: "Anthropic Agent SDK", bold: true },
        ") that fixes bugs, adds features, and summarizes discussions on request."
      ]),

      // -- SDE --
      titleLine("Software Development Engineer"),
      companyDateLine("Criteo", "Ann Arbor, MI", "Oct 2023 - Oct 2024"),
      bullet([
        "Shipped auction-based bidding and multi-creative display ads, ",
        { text: "giving advertisers new ways to buy and scale campaigns", bold: true },
        " (",
        { text: "C#/.NET 6-8, Angular", bold: true },
        ")."
      ]),
      bullet([
        "Set up Grafana dashboards, alerting, and operational docs for Hermes, the mission-critical AWS messaging system I later became the go-to leader for, ",
        { text: "transforming incident response", bold: true },
        "."
      ]),
      bullet([
        "Fixed critical ad-serving production bugs that risked ",
        { text: "significant financial impact", bold: true },
        "."
      ]),

      // -- Rocket --
      titleLine("Software Developer Co-op (2x)"),
      companyDateLine("Rocket Innovation Studio", "Windsor, ON", "May - Aug 2021, Jan - May 2022"),
      bullet([
        "Delivered ",
        { text: "fully unit-tested", bold: true },
        " Angular 11/AngularJS components and .NET Core 3/5/6 backend services."
      ]),

      // -- TA & Research --
      titleLine("Head Teaching Assistant & Lead Research Assistant"),
      companyDateLine("University of Windsor", "Windsor, ON", "Sep 2021 - Aug 2023"),
      bullet([
        "Led connected-vehicle ",
        { text: "ML research", bold: true },
        ", managing and mentoring a team of research assistants."
      ]),
      bullet([
        "Published a cognitive load recognition paper (scikit-learn, eye-tracking) at ",
        { text: "ACR’23 (Springer)", bold: true },
        ". ",
        link("Cognitive Load Recognition via Eye-Tracking", "https://doi.org/10.1007/978-3-031-33743-7_7"),
      ]),
      bullet([
        "Engineered a configurable ML pipeline (scikit-learn, TensorFlow, Pandas) that classifies connected-vehicle safety messages. ",
        link("ConnectedDrivingPipelineV4", "https://www.aaroncollins.info/ConnectedDrivingPipelineV4/"),
      ]),
      bullet([
        { text: "Head TA", bold: true },
        " for the Python programming course, plus TA for one other course."
      ]),

      // ═══ PROJECTS (page 2) ═══
      new Paragraph({
        children: [new (require("docx").PageBreak)()],
      }),
      sectionHeading("Projects"),

      projectTitleLine("Finledger", "2026 - Present"),
      bullet([
        "Building a finance app that detects recurring costs to ",
        { text: "optimize monthly spending", bold: true },
        ", with ",
        { text: "on-device AI", bold: true },
        " and password-protected iCloud sync; ships to ",
        { text: "6 platforms from one Flutter codebase", bold: true },
        ". ",
        link("Site", "https://www.aaroncollins.info/finledger-site/"),
      ]),

      projectTitleLine("SOPHIE", "2025 - 2026"),
      bullet([
        "Personal AI assistant on Slack, Discord, or phone that ",
        { text: "works autonomously", bold: true },
        ": manages email, preps tax documents, makes calls, and builds side projects (MCP, skills). ",
        link("sophie", "https://www.aaroncollins.info/projects/sophie.html"),
        " | ",
        link("sophie-mcp", "https://www.aaroncollins.info/projects/sophie-mcp.html"),
        " | ",
        link("sophie-voice-bridge", "https://www.aaroncollins.info/projects/sophie-voice-bridge.html"),
      ]),

      projectTitleLine("PortableRalph", "2026"),
      bullet([
        "Open-source autonomous AI coding loop for any repo on Linux, macOS, and Windows; ",
        { text: "13+ GitHub stars", bold: true },
        ". ",
        link("GitHub", "https://github.com/aaron777collins/portableralph"),
        " | ",
        link("Docs", "https://www.portableralph.com/"),
      ]),

      projectTitleLine("ZeroTier-GUI-Arch", "2024 - 2025"),
      bullet([
        "Ported and improved ZeroTier-GUI for Steam Deck and Arch Linux, ",
        { text: "making private networking easy for non-technical users", bold: true },
        "; ",
        { text: "41+ GitHub stars", bold: true },
        ", released via Flatpak. ",
        link("GitHub", "https://github.com/aaron777collins/ZeroTier-GUI-Arch"),
        " | ",
        link("Docs", "https://www.aaroncollins.info/projects/ZeroTier-GUI-Arch.html"),
      ]),

      projectTitleLine("Minecraft Datapack Language (MDL)", "2025"),
      bullet([
        "Programming language and compiler that brings if/else, loops, and variables to Minecraft datapacks; ",
        { text: "published on PyPI", bold: true },
        " with a VS Code extension. ",
        link("GitHub", "https://github.com/aaron777collins/MinecraftDatapackLanguage"),
        " | ",
        link("Docs", "https://www.aaroncollins.info/MinecraftDatapackLanguage/"),
      ]),

      projectTitleLine("Control Windows", "Jun 2020 - Jan 2023"),
      bullet([
        "Desktop automation tool (Python, PyAutoGUI) that auto-generates automation code. ",
        link("GitHub", "https://github.com/aaron777collins/Control-Windows"),
        " | ",
        link("Docs", "https://www.aaroncollins.info/projects/Control-Windows.html"),
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
        "C#, Python, TypeScript, JavaScript, Rust, Java, C, R, SQL, Bash"
      ]),
      bullet([
        { text: "AI & ML: ", bold: true },
        "AI agents, multi-agent orchestration, Anthropic Agent SDK, Claude Code, MCP, Cursor, AMP, Ollama, TensorFlow, scikit-learn, Pandas, NumPy"
      ]),
      bullet([
        { text: "Ad Tech: ", bold: true },
        "Retail media, onsite display ads, real-time ad auctions, ad validation, ad delivery forecasting"
      ]),
      bullet([
        { text: "Frameworks & Web: ", bold: true },
        ".NET (Core 3-10), Angular, React, Flask, Node.js, HTML, CSS/SCSS"
      ]),
      bullet([
        { text: "Infrastructure & DevOps: ", bold: true },
        "AWS (SQS, Lambda, ECS), ScyllaDB, Docker, Jenkins, Grafana, Kibana, NGINX, Caddy, Authelia (SSO), CI/CD"
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Resume_Aaron_Collins.docx", buffer);
  console.log("Done");
});
