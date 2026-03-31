const pptxgen = require("pptxgenjs");
const path = require("path");

const OUT_DIR = "/Users/haoxingwang/WorkBuddy/QQx小游戏品牌方案/03_PPT营销案模板";

// ===== Color Palettes =====
const PALETTES = {
  universal: { // QQ品牌蓝
    primary: "1A73E8",
    dark: "0D47A1",
    light: "E8F0FE",
    accent: "FF6D00",
    text: "212121",
    subtext: "666666",
    white: "FFFFFF",
    bg: "F5F7FA",
    bgDark: "1A2742",
  },
  casual: { // 休闲游戏 - 明亮活泼
    primary: "FF6B6B",
    dark: "C92A2A",
    light: "FFF0F0",
    accent: "FFD93D",
    text: "2D3436",
    subtext: "636E72",
    white: "FFFFFF",
    bg: "FFF9F0",
    bgDark: "2D1B69",
  },
  slg: { // SLG - 深色科技
    primary: "00B4D8",
    dark: "0A1628",
    light: "CAF0F8",
    accent: "FF6B35",
    text: "E2E8F0",
    subtext: "94A3B8",
    white: "FFFFFF",
    bg: "0F172A",
    bgDark: "020617",
  },
  sim: { // 模拟经营 - 温暖自然
    primary: "2D9F5D",
    dark: "1A5632",
    light: "ECFDF5",
    accent: "F59E0B",
    text: "1F2937",
    subtext: "6B7280",
    white: "FFFFFF",
    bg: "F0FDF4",
    bgDark: "14532D",
  },
};

function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.12 };
}

// ===== Slide Builders =====

function addCoverSlide(pres, p, title, subtitle) {
  const slide = pres.addSlide();
  slide.background = { color: p.bgDark };

  // Decorative shapes
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: p.primary } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.8, w: 10, h: 0.06, fill: { color: p.primary, transparency: 40 } });

  // Title
  slide.addText(title, {
    x: 0.8, y: 1.2, w: 8.5, h: 1.6,
    fontSize: 36, fontFace: "Arial Black", color: p.white, bold: true, margin: 0,
  });

  // Subtitle
  slide.addText(subtitle, {
    x: 0.8, y: 3.0, w: 7, h: 0.8,
    fontSize: 16, fontFace: "Arial", color: p.subtext, margin: 0,
  });

  // Footer
  slide.addText("CONFIDENTIAL | QQ Media × Mini Game", {
    x: 0.8, y: 5.0, w: 5, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: p.subtext, margin: 0,
  });
}

function addSectionSlide(pres, p, sectionNum, sectionTitle, sectionDesc) {
  const slide = pres.addSlide();
  slide.background = { color: p.primary };

  slide.addText(`0${sectionNum}`, {
    x: 0.8, y: 1.0, w: 2, h: 1.5,
    fontSize: 72, fontFace: "Arial Black", color: p.white, bold: true, margin: 0,
    transparency: 20,
  });

  slide.addText(sectionTitle, {
    x: 0.8, y: 2.5, w: 8, h: 1.0,
    fontSize: 32, fontFace: "Arial", color: p.white, bold: true, margin: 0,
  });

  slide.addText(sectionDesc, {
    x: 0.8, y: 3.6, w: 7, h: 0.6,
    fontSize: 14, fontFace: "Arial", color: p.white, margin: 0, transparency: 30,
  });
}

function addContentSlide(pres, p, title, items, isDark = false) {
  const slide = pres.addSlide();
  slide.background = { color: isDark ? p.bgDark : p.bg };
  const textColor = isDark ? p.white : p.text;
  const subtextColor = isDark ? p.subtext : p.subtext;

  // Title
  slide.addText(title, {
    x: 0.6, y: 0.3, w: 9, h: 0.7,
    fontSize: 24, fontFace: "Arial", color: textColor, bold: true, margin: 0,
  });
  // Title underline
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.0, w: 1.5, h: 0.05, fill: { color: p.primary } });

  // Content cards (2-column grid)
  const cols = 2;
  const cardW = 4.2;
  const cardH = 1.8;
  const startX = 0.6;
  const startY = 1.3;
  const gapX = 0.5;
  const gapY = 0.3;

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: isDark ? "1E293B" : p.white },
      shadow: makeShadow(),
    });

    // Accent bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH,
      fill: { color: p.primary },
    });

    slide.addText(item.title, {
      x: x + 0.25, y: y + 0.15, w: cardW - 0.4, h: 0.5,
      fontSize: 14, fontFace: "Arial", color: textColor, bold: true, margin: 0,
    });

    slide.addText(item.desc, {
      x: x + 0.25, y: y + 0.6, w: cardW - 0.4, h: 1.0,
      fontSize: 11, fontFace: "Arial", color: subtextColor, margin: 0,
    });
  });
}

function addDataSlide(pres, p, title, stats) {
  const slide = pres.addSlide();
  slide.background = { color: p.bgDark };

  slide.addText(title, {
    x: 0.6, y: 0.3, w: 9, h: 0.7,
    fontSize: 24, fontFace: "Arial", color: p.white, bold: true, margin: 0,
  });

  const cardW = (9 - (stats.length - 1) * 0.3) / stats.length;
  stats.forEach((stat, i) => {
    const x = 0.6 + i * (cardW + 0.3);
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.3, w: cardW, h: 3.5,
      fill: { color: "1E293B" },
      shadow: makeShadow(),
    });

    slide.addText(stat.value, {
      x, y: 1.5, w: cardW, h: 1.2,
      fontSize: 44, fontFace: "Arial Black", color: p.primary, bold: true, align: "center", margin: 0,
    });

    slide.addText(stat.label, {
      x, y: 2.7, w: cardW, h: 0.5,
      fontSize: 14, fontFace: "Arial", color: p.white, bold: true, align: "center", margin: 0,
    });

    slide.addText(stat.desc, {
      x: x + 0.2, y: 3.3, w: cardW - 0.4, h: 1.2,
      fontSize: 10, fontFace: "Arial", color: p.subtext, align: "center", margin: 0,
    });
  });
}

function addTableSlide(pres, p, title, headers, rows) {
  const slide = pres.addSlide();
  slide.background = { color: p.bg };

  slide.addText(title, {
    x: 0.6, y: 0.3, w: 9, h: 0.6,
    fontSize: 22, fontFace: "Arial", color: p.text, bold: true, margin: 0,
  });

  const headerRow = headers.map(h => ({
    text: h, options: { fill: { color: p.primary }, color: "FFFFFF", bold: true, fontSize: 11, fontFace: "Arial" }
  }));

  const dataRows = rows.map((row, idx) =>
    row.map(cell => ({
      text: cell, options: {
        fill: { color: idx % 2 === 0 ? p.white : p.light },
        color: p.text, fontSize: 10, fontFace: "Arial"
      }
    }))
  );

  slide.addTable([headerRow, ...dataRows], {
    x: 0.5, y: 1.1, w: 9,
    border: { pt: 0.5, color: "D9D9D9" },
    colW: Array(headers.length).fill(9 / headers.length),
    rowH: 0.4,
  });
}

function addClosingSlide(pres, p, mainText, contactInfo) {
  const slide = pres.addSlide();
  slide.background = { color: p.bgDark };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: p.primary } });

  slide.addText(mainText, {
    x: 1, y: 1.5, w: 8, h: 1.5,
    fontSize: 32, fontFace: "Arial Black", color: p.white, bold: true, align: "center", margin: 0,
  });

  slide.addText(contactInfo, {
    x: 1.5, y: 3.5, w: 7, h: 1.0,
    fontSize: 14, fontFace: "Arial", color: p.subtext, align: "center", margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.425, w: 10, h: 0.2, fill: { color: p.primary } });
}

// ===== Create Presentations =====

async function createUniversalPPT() {
  const pres = new pptxgen();
  const p = PALETTES.universal;
  pres.layout = "LAYOUT_16x9";
  pres.author = "QQ Media";
  pres.title = "QQ Media x Mini Game Marketing Proposal";

  addCoverSlide(pres, p,
    "QQ媒体 × 微信小游戏\n品牌营销合作方案",
    "[客户名称] 专属定制 ｜ 2026年Q2"
  );

  addSectionSlide(pres, p, 1, "QQ媒体生态价值", "触达中国最年轻的社交用户群体");

  addDataSlide(pres, p, "QQ核心数据一览", [
    { value: "5亿+", label: "月活跃用户", desc: "中国最大的年轻人社交平台之一" },
    { value: "70%", label: "25岁以下用户占比", desc: "Z世代核心社交阵地" },
    { value: "高活跃", label: "社交互动频次", desc: "熟人关系链驱动的深度社交" },
  ]);

  addContentSlide(pres, p, "QQ媒体独特优势", [
    { title: "最年轻的社交用户群", desc: "QQ是中国互联网唯一以15-25岁为核心的社交平台，Z世代浓度极高" },
    { title: "深度熟人关系链", desc: "同学/朋友的强关系推荐，信任度远超算法推荐和KOL种草" },
    { title: "天然游戏基因", desc: "QQ游戏大厅、QQ空间小游戏等积淀，用户对'在QQ上玩游戏'有天然心智" },
    { title: "差异化社交场景", desc: "聊天/空间/音乐/阅读等多元场景，'一起玩'而非'刷到就玩'" },
  ]);

  addSectionSlide(pres, p, 2, "用户画像与数据洞察", "精准匹配您的目标用户");

  addContentSlide(pres, p, "QQ × [游戏名] 用户画像匹配", [
    { title: "年龄画像", desc: "QQ核心用户15-25岁，与[游戏名]目标用户高度重合\n[此处填入具体重合度数据]" },
    { title: "兴趣标签", desc: "二次元/潮流文化/社交游戏/音乐\n[此处填入与游戏品类匹配的兴趣标签]" },
    { title: "社交行为", desc: "日均社交互动频次高\n好友关系链活跃，利于社交裂变传播" },
    { title: "地域分布", desc: "一二线城市+校园用户为主\n[此处填入城市分布数据]" },
  ]);

  addSectionSlide(pres, p, 3, "合作模式与资源矩阵", "品效协同的营销解决方案");

  addTableSlide(pres, p, "QQ媒体营销资源矩阵", 
    ["资源类型", "形式", "核心价值", "参考价格"],
    [
      ["开屏广告", "全屏展示（图片/视频）", "超强品牌曝光", "CPT定价"],
      ["信息流广告", "QQ空间/看点原生", "融入社交场景", "CPM/CPC"],
      ["社交互动广告", "好友组队/挑战赛", "社交裂变传播", "定制报价"],
      ["QQ音乐联动", "歌单/BGM合作", "沉浸式品牌体验", "定制报价"],
      ["品牌专区", "定制化品牌空间", "深度品牌展示", "定制报价"],
    ]
  );

  addContentSlide(pres, p, "创意营销玩法推荐", [
    { title: "社交挑战赛", desc: "发起好友挑战，利用关系链驱动病毒传播\n预估参与人次：XX万+" },
    { title: "QQ空间互动广告", desc: "在QQ空间信息流中嵌入可互动的小游戏体验\n支持试玩+一键跳转" },
    { title: "好友组队活动", desc: "邀请QQ好友组队完成游戏任务\n社交绑定提升留存和活跃" },
    { title: "IP联名合作", desc: "与QQ自有IP（如QQ企鹅）联名\n提升品牌年轻化认知" },
  ], true);

  addSectionSlide(pres, p, 4, "成功案例", "验证过的合作方法论");

  addContentSlide(pres, p, "案例展示：[案例名称]", [
    { title: "合作背景", desc: "[客户名称] × QQ媒体\n品类：[游戏品类]  |  档期：[投放时间]  |  预算：[金额]" },
    { title: "营销策略", desc: "[此处填入案例的核心营销策略和资源组合]" },
    { title: "效果数据", desc: "总曝光：XX万  |  点击率：X.X%\n转化量：XX万  |  ROI：X.X" },
    { title: "客户评价", desc: "[此处填入客户的正面反馈或评价]" },
  ]);

  addSectionSlide(pres, p, 5, "合作报价与排期", "灵活的合作方案");

  addTableSlide(pres, p, "推荐合作套餐", 
    ["套餐", "资源组合", "预估效果", "合作价格"],
    [
      ["基础版", "信息流广告 × 2周", "曝光XX万+，点击XX万+", "[价格]"],
      ["进阶版", "信息流 + 社交互动 × 1个月", "曝光XX万+，互动XX万+", "[价格]"],
      ["旗舰版", "开屏 + 信息流 + 社交挑战赛 + IP联名", "全链路品效协同", "[价格]"],
    ]
  );

  addClosingSlide(pres, p,
    "期待与您携手共创\n社交营销新可能",
    "[销售姓名]  |  [手机号]  |  [企微二维码]\nQQ媒体 × 微信小游戏 商业化团队"
  );

  await pres.writeFile({ fileName: path.join(OUT_DIR, "QQx小游戏通用营销方案模板.pptx") });
  console.log("✅ 通用营销方案模板.pptx");
}

async function createCasualPPT() {
  const pres = new pptxgen();
  const p = PALETTES.casual;
  pres.layout = "LAYOUT_16x9";
  pres.author = "QQ Media";
  pres.title = "QQ x Casual Game Marketing Proposal";

  addCoverSlide(pres, p,
    "QQ媒体 × 休闲小游戏\n品牌营销定制方案",
    "[游戏名称] 专属方案 ｜ 轻松有趣，社交无限"
  );

  addSectionSlide(pres, p, 1, "休闲游戏行业趋势", "小游戏赛道的黄金时代");

  addDataSlide(pres, p, "休闲小游戏行业核心趋势", [
    { value: "爆发式", label: "市场增长", desc: "微信小游戏DAU持续攀升\n休闲品类占比最高" },
    { value: "IAA为主", label: "变现模式", desc: "广告变现为核心\n品牌广告需求旺盛" },
    { value: "社交化", label: "玩法趋势", desc: "社交裂变成为核心增长引擎\n好友互动提升留存" },
  ]);

  addContentSlide(pres, p, "休闲游戏 × QQ：天然契合", [
    { title: "用户画像完美匹配", desc: "休闲游戏核心用户15-30岁，与QQ年轻用户高度重合\n碎片化社交场景 = 休闲游戏最佳入口" },
    { title: "轻互动创意玩法", desc: "QQ聊天场景中嵌入小游戏互动\n'发一局'比'发表情'更有趣" },
    { title: "社交裂变天然优势", desc: "休闲游戏的分享意愿最强\nQQ好友链驱动的社交裂变效率最高" },
    { title: "品牌年轻化利器", desc: "休闲+社交+QQ = 品牌年轻化最佳组合\n轻松有趣的品牌形象塑造" },
  ]);

  addSectionSlide(pres, p, 2, "定制化合作方案", "为您的游戏量身打造");

  addContentSlide(pres, p, "推荐创意玩法", [
    { title: "好友PK挑战赛", desc: "邀请好友PK分数，排行榜激发竞争欲\n预估参与：XX万人次" },
    { title: "QQ表情联动", desc: "游戏角色制作为QQ表情包\n社交传播+品牌曝光双赢" },
    { title: "聊天小窗试玩", desc: "聊天界面直接嵌入游戏试玩\n'发一局给你'的社交新体验" },
    { title: "节日限定活动", desc: "结合节日热点推出限定活动\n短周期爆发式流量" },
  ], true);

  addTableSlide(pres, p, "效果预估（短周期投放）",
    ["指标", "基础版(2周)", "进阶版(1个月)", "备注"],
    [
      ["总曝光", "XX万+", "XX万+", "QQ信息流 + 空间"],
      ["总点击", "XX万+", "XX万+", "含试玩点击"],
      ["社交传播人次", "XX万+", "XX万+", "好友裂变"],
      ["预估CTR", "X.X%+", "X.X%+", "高于行业均值"],
      ["预估CPA", "XX元", "XX元", "休闲品类优化"],
    ]
  );

  addClosingSlide(pres, p,
    "让每一次社交互动\n都成为一场轻松游戏",
    "[销售姓名]  |  [手机号]  |  [企微二维码]\nQQ媒体 × 休闲小游戏 定制方案"
  );

  await pres.writeFile({ fileName: path.join(OUT_DIR, "休闲游戏品类定制方案模板.pptx") });
  console.log("✅ 休闲游戏品类定制方案模板.pptx");
}

async function createSLGPPT() {
  const pres = new pptxgen();
  const p = PALETTES.slg;
  pres.layout = "LAYOUT_16x9";
  pres.author = "QQ Media";
  pres.title = "QQ x SLG Marketing Proposal";

  addCoverSlide(pres, p,
    "QQ媒体 × SLG小游戏\n品牌营销定制方案",
    "[游戏名称] 专属方案 ｜ 策略对决，社交征服"
  );

  addSectionSlide(pres, p, 1, "SLG行业趋势洞察", "重度赛道的小游戏突围之路");

  addDataSlide(pres, p, "SLG小游戏行业核心数据", [
    { value: "手转小", label: "核心趋势", desc: "传统SLG以小游戏形态\n降低用户体验门槛" },
    { value: "高LTV", label: "用户价值", desc: "SLG用户付费意愿强\n长线留存表现优秀" },
    { value: "社交强", label: "玩法特征", desc: "联盟/对抗/外交\n社交绑定是留存核心" },
  ]);

  addContentSlide(pres, p, "SLG × QQ：社交关系链的战略价值", [
    { title: "硬核玩家社交需求", desc: "SLG玩家需要稳定的社交关系来支撑联盟玩法\nQQ的强关系链是天然的联盟基础" },
    { title: "QQ群 = 天然盟会", desc: "SLG玩家习惯用QQ群进行战术沟通\nQQ是SLG社交的'第二战场'" },
    { title: "长线用户培育", desc: "SLG需要长线用户运营\nQQ的社交场景支持持续触达" },
    { title: "年轻用户拓新", desc: "SLG需要拓展年轻用户\nQQ是触达Z世代SLG潜在玩家的最佳渠道" },
  ], true);

  addSectionSlide(pres, p, 2, "深度互动合作方案", "策略+社交的差异化营销");

  addContentSlide(pres, p, "推荐创意玩法", [
    { title: "QQ群服联盟战", desc: "通过QQ群组建联盟，发起跨服对抗\n社交绑定 + 竞技激情 = 超高活跃" },
    { title: "策略挑战赛", desc: "在QQ空间发起策略脑力挑战\n展示SLG策略深度，吸引硬核玩家" },
    { title: "指挥官直播", desc: "QQ直播间进行大型会战直播\n观众实时参与指挥决策" },
    { title: "QQ游戏中心专题", desc: "QQ游戏中心开设SLG专题页\n深度展示游戏世界观和玩法" },
  ]);

  addTableSlide(pres, p, "效果预估（长线投放）",
    ["指标", "首月(品牌期)", "2-3月(效果期)", "备注"],
    [
      ["总曝光", "XX万+", "XX万+", "品牌+效果全覆盖"],
      ["核心用户获取", "XX万+", "XX万+", "精准SLG用户"],
      ["7日留存率", "XX%+", "XX%+", "社交绑定提升留存"],
      ["联盟创建数", "XX个+", "XX个+", "QQ群驱动"],
      ["预估LTV", "XX元", "XX元", "长线价值"],
    ]
  );

  addClosingSlide(pres, p,
    "用社交关系链\n构建游戏世界的战略联盟",
    "[销售姓名]  |  [手机号]  |  [企微二维码]\nQQ媒体 × SLG小游戏 定制方案"
  );

  await pres.writeFile({ fileName: path.join(OUT_DIR, "SLG游戏品类定制方案模板.pptx") });
  console.log("✅ SLG游戏品类定制方案模板.pptx");
}

async function createSimPPT() {
  const pres = new pptxgen();
  const p = PALETTES.sim;
  pres.layout = "LAYOUT_16x9";
  pres.author = "QQ Media";
  pres.title = "QQ x Simulation Game Marketing Proposal";

  addCoverSlide(pres, p,
    "QQ媒体 × 模拟经营小游戏\n品牌营销定制方案",
    "[游戏名称] 专属方案 ｜ 温暖经营，好友共建"
  );

  addSectionSlide(pres, p, 1, "模拟经营行业趋势", "治愈系游戏的崛起");

  addDataSlide(pres, p, "模拟经营品类核心趋势", [
    { value: "高增长", label: "品类增速", desc: "模拟经营品类在小游戏生态中\n增速位居前三" },
    { value: "泛休闲", label: "用户画像", desc: "女性用户占比高\n18-30岁年轻用户为主" },
    { value: "中留存", label: "留存表现", desc: "中周期留存表现优秀\n用户粘性高于纯休闲" },
  ]);

  addContentSlide(pres, p, "模拟经营 × QQ：社交共建的温暖体验", [
    { title: "好友互访互助", desc: "模拟经营天然适合好友互动\n'来我农场帮忙浇水' = 社交+游戏结合" },
    { title: "QQ空间种草", desc: "在QQ空间分享我的'小店/农场/城市'\n治愈系内容自然传播" },
    { title: "女性用户触达", desc: "QQ的年轻女性用户群\n与模拟经营品类高度匹配" },
    { title: "沉浸式品牌植入", desc: "在游戏场景中植入品牌元素\n自然融入、用户体验友好" },
  ]);

  addSectionSlide(pres, p, 2, "沉浸式合作方案", "在温暖的游戏世界里讲品牌故事");

  addContentSlide(pres, p, "推荐创意玩法", [
    { title: "好友农场/小镇互访", desc: "在QQ内直接访问好友的游戏世界\n'来我的小镇逛逛' = 最治愈的社交邀请" },
    { title: "经营日记分享", desc: "自动生成精美的经营日记\n一键分享到QQ空间/朋友说说" },
    { title: "联名主题装扮", desc: "QQ × 游戏联名限定装扮/皮肤\n社交身份认同 + 收藏欲" },
    { title: "节日合作经营", desc: "节日限定经营活动\n好友协力完成特殊任务" },
  ], true);

  addTableSlide(pres, p, "效果预估（中周期投放）",
    ["指标", "首月", "2-3月", "备注"],
    [
      ["总曝光", "XX万+", "XX万+", "QQ空间 + 信息流"],
      ["试玩转化", "XX万+", "XX万+", "含互访体验"],
      ["社交分享次数", "XX万+", "XX万+", "经营日记分享"],
      ["30日留存", "XX%+", "XX%+", "社交绑定提升"],
      ["品牌好感度", "+XX%", "+XX%", "沉浸式植入效果"],
    ]
  );

  addClosingSlide(pres, p,
    "在QQ的社交世界里\n一起建造属于我们的温暖故事",
    "[销售姓名]  |  [手机号]  |  [企微二维码]\nQQ媒体 × 模拟经营小游戏 定制方案"
  );

  await pres.writeFile({ fileName: path.join(OUT_DIR, "模拟经营品类定制方案模板.pptx") });
  console.log("✅ 模拟经营品类定制方案模板.pptx");
}

// ===== Main =====
async function main() {
  await createUniversalPPT();
  await createCasualPPT();
  await createSLGPPT();
  await createSimPPT();
  console.log("\n🎉 所有4份PPT模板创建完成！");
}

main().catch(console.error);
