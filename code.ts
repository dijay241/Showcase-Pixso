// code.ts
// TS таргет: ES2019. Собираем из выделенных фреймов плоский список объектов
// (TEXT / SVG / IMAGE) с абсолютными координатами, размерами и поворотом.

  type ExportTextRun = {
    text: string;
    options: {
      fontFace?: string;
      fontSize?: number;
      color?: string;        // hex без '#'
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strike?: "sngStrike" | "dblStrike";
      charSpacing?: number;  // в пунктах
      lineSpacing?: number;  // в пунктах
      breakLine?: boolean;
      bullet?: boolean | { type?: string; code?: string; style?: string };
      indentLevel?: number;
    };
  };
  
  type ExportItemBase = {
    id: string;
    type: "TEXT" | "SVG" | "IMAGE" | "RECT" | "LINE" | "GROUP";
    xIn?: number; // inches
    yIn?: number;
    wIn?: number;
    hIn?: number;
    rotate?: number; // degrees
    name?: string;
    zIndex: number;
  };
  
  type ExportTextItem = ExportItemBase & {
    type: "TEXT";
    runs: ExportTextRun[];
    valign?: "top" | "middle" | "bottom";
    align?: "left" | "center" | "right";
    boxWIn: number;
    boxHIn: number;
    textwIn?: number;
  };
  
  type ExportSvgItem = ExportItemBase & {
    type: "SVG";
    svgBase64: string; // "image/svg+xml;base64,...."
    strokeWeight?: number;
  };
  
  type ExportImageItem = ExportItemBase & {
    type: "IMAGE";
    mime: "image/png" | "image/jpeg";
    dataBase64: string; // "image/png;base64,...."
  };

  type ExportRectItem = ExportItemBase & {
    type: "RECT";
    fillHex: string;
    borderHex?: string;
    borderWidth?: number;
    borderStyle?: string;
  };

  type ExportLineItem = ExportItemBase & {
    type: "LINE";
    strokeHex: string;
    strokePt: number; // thickness in points
    cap?: "round" | "square" | "flat";
    dash?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot";
    beginArrow?: boolean;
    endArrow?: boolean;
  };

  type ExportGroupItem = ExportItemBase & {
    type: "GROUP";
    children: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem>;
  };
  
  type ExportFrame = {
    id: string;
    name: string;
    widthPx: number;
    heightPx: number;
    slideWIn: number;  // = widthPx / 96
    slideHIn: number;  // = heightPx / 96
    bgColorHex?: string;
    items: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem>;
  };
  
  figma.on("run", async ({ command }) => {
    const selection = figma.currentPage.selection;
    const frames = selection.filter(n =>
      n.type === "FRAME" || n.type === "COMPONENT" || n.type === "INSTANCE"
    ) as (FrameNode | ComponentNode | InstanceNode)[];
  
    if (!frames.length) {
      figma.notify("Select at least one frame");
      figma.closePlugin();
      return;
    }
  
    const exportFrames: ExportFrame[] = [];
    let zCounter = 0;
  
    for (const frame of frames) {
      const { x: frameAbsX, y: frameAbsY } = absOrigin(frame);
      const slideWIn = pxToIn(frame.width);
      const slideHIn = pxToIn(frame.height);
      const bgHex = frameBackgroundHex(frame);
  
      const items: Array<ExportTextItem | ExportSvgItem | ExportImageItem | ExportRectItem | ExportLineItem | ExportGroupItem> = [];
      // Рекурсивный обход: сначала фоны фреймов (как прямоугольники), затем дети, затем прочие элементы
      const walk = async (node: SceneNode, parentGroup?: ExportGroupItem) => {
        if (!node.visible) return;
        const { x: absX, y: absY } = absOrigin(node as SceneNode);
        const xIn = pxToIn(absX - frameAbsX);
        const yIn = pxToIn(absY - frameAbsY);
        // Получаем поворот из свойства rotation или вычисляем из трансформации
        let rotate = 0;
        if ('rotation' in node && typeof (node as any).rotation === 'number') {
          rotate = (node as any).rotation;
        } else {
          rotate = rotationFromTransform((node as SceneNode).absoluteTransform);
        }
        
        if (rotate !== 0) {
          console.log('Node:', node.name, 'Type:', node.type, 'Rotation:', rotate, 'Source:', 'rotation' in node ? 'property' : 'transform');
        }
        const wIn = hasSize(node) ? pxToIn(node.width) : undefined;
        const hIn = hasSize(node) ? pxToIn(node.height) : undefined;

        // Фон и рамка вложенных фреймов/компонентов/инстансов
        if (node.type === "FRAME" || node.type === "COMPONENT" || node.type === "INSTANCE") {
          const bg = frameBackgroundHex(node as FrameNode | ComponentNode | InstanceNode);
          const border = frameBorder(node as FrameNode | ComponentNode | InstanceNode);
          
          if ((bg || border) && hasSize(node) && (node.width > 0 || node.height > 0)) {
            items.push({
              id: node.id + "::frame",
              type: "RECT",
              xIn, yIn,
              wIn, hIn,
              rotate,
              fillHex: bg || "00000000", // фон или прозрачность
              borderHex: border ? border.color : undefined,
              borderWidth: border ? border.width : undefined,
              borderStyle: border ? border.style : undefined,
              name: node.name + " (frame)",
              zIndex: zCounter++
            });
          }
        }

        if ("children" in node && (node as any).children) {
          // Если это группа, создаём элемент группы
          if (node.type === "GROUP") {
            const groupItem: ExportGroupItem = {
              id: node.id,
              type: "GROUP",
              xIn, yIn, wIn, hIn, rotate,
              name: node.name,
              zIndex: zCounter++,
              children: []
            };
            
            // Рекурсивно обрабатываем детей группы
            for (const child of (node as any).children as SceneNode[]) {
              await walk(child, groupItem);
            }
            
            // Добавляем группу в родительский контейнер
            if (parentGroup) {
              parentGroup.children.push(groupItem);
            } else {
              items.push(groupItem);
            }
          } else {
            // Для фреймов и других контейнеров просто рекурсивно обходим детей
            for (const child of (node as any).children as SceneNode[]) {
              await walk(child, parentGroup);
            }
          }
        }

        if (node.type === "TEXT") {
          const textNode = node as TextNode;
          const runs: ExportTextRun[] = [];
  
          const segments = textNode.getStyledTextSegments(["fontName", "fontSize", "fills", "textDecoration", "letterSpacing", "lineHeight"]);
          
          // Получаем информацию о списках из Figma
          const listOptions = (textNode as any).getRangeListOptions ? (textNode as any).getRangeListOptions(0, textNode.characters.length) : null;
          
          for (const seg of segments) {
            const str = textNode.characters.substring(seg.start, seg.end);
            if (!str.length) continue;

            const fill = firstSolidPaint(seg.fills);
            const colorHex = fill ? rgbToHex(fill.color, fill.opacity) : "000000";
            const underline = seg.textDecoration === "UNDERLINE";
            const strike = seg.textDecoration === "STRIKETHROUGH" ? "sngStrike" : undefined;
            const isBold = typeof seg.fontName === 'object' && isFont(seg.fontName)
              ? /(bold|semibold|demi|medium)/i.test(seg.fontName.style || '')
              : false;
            const isItalic = typeof seg.fontName === 'object' && isFont(seg.fontName)
              ? /italic/i.test(seg.fontName.style || '')
              : false;

            // letterSpacing: Figma -> pptx (pt)
            let charSpacingPt: number | undefined;
            if (seg.letterSpacing) {
              if (seg.letterSpacing.unit === "PIXELS")
                charSpacingPt = pxToPt(seg.letterSpacing.value);
              if (seg.letterSpacing.unit === "PERCENT")
                charSpacingPt = (seg.fontSize || 12) * pxToPt(seg.letterSpacing.value / 100);
            }

            // lineHeight: либо pts, либо multiple
            let lineSpacing: number | undefined;
            if (seg.lineHeight) {
              if (seg.lineHeight.unit === "PIXELS")
                lineSpacing = pxToPt(seg.lineHeight.value);
              if (seg.lineHeight.unit === "PERCENT")
                lineSpacing = (seg.fontSize || 12) * pxToPt(seg.lineHeight.value / 100);
            }

            // Определяем свойства списка из Figma
            let bullet: boolean | { type?: string; code?: string; style?: string } | undefined;
            let indentLevel: number | undefined;
            
            if (listOptions && Object.keys(listOptions).length) {
            
              const listType = listOptions.type;
              
              if (listType === 'ORDERED') {
                bullet = { type: 'number' };
              } else if (listType === 'UNORDERED') {
                bullet = true;
              } else if (listType === 'NONE') {
                bullet = false;
              } else {
                bullet = false;
              }
              
            }

            runs.push({
              text: str,
              options: {
                fontFace: isFont(seg.fontName) ? seg.fontName.family : undefined,
                fontSize: typeof seg.fontSize === 'number' ? Math.round(pxToPt(seg.fontSize)) : undefined,
                color: colorHex,
                underline,
                strike,
                bold: isBold || undefined,
                italic: isItalic || undefined,
                charSpacing: charSpacingPt,
                lineSpacing,
                bullet
              }
            });
          }
  
          const align = mapAlign(textNode.textAlignHorizontal);
          const valign = mapVAlign(textNode.textAlignVertical);
          const textwIn = wIn ? wIn * 1.05 : undefined;
  
          const textItem: ExportTextItem = {
            id: textNode.id,
            type: "TEXT",
            xIn, yIn,
            textwIn, hIn,
            boxWIn: textwIn || 1,
            boxHIn: hIn || 1,
            rotate,
            runs,
            align, valign,
            name: textNode.name,
            zIndex: zCounter++
          };
          
          // Добавляем в родительскую группу или в основной список
          if (parentGroup) {
            parentGroup.children.push(textItem);
          } else {
            items.push(textItem);
          }
        }
        else if (isVectorNode(node)) {
          if (!hasRenderableSize(node)) return;
          try {
            const data = await (node as unknown as ExportMixin).exportAsync({
              format: "SVG",
              svgOutlineText: false, // важно: не превращать текст в кривые
              svgIdAttribute: false,
              svgSimplifyStroke: true
            });
            
            const svgBase64 = "data:image/svg+xml;base64," + figma.base64Encode(data);
            const strokeWeight = (node as any).strokeWeight;

            const svgItem: ExportSvgItem = {
              id: node.id,
              type: "SVG",
              xIn, yIn, wIn, hIn: hIn || (strokeWeight ? strokeWeight * 0.01 : 0.01),
              rotate,
              svgBase64,
              name: node.name,
              zIndex: zCounter++
            };
            
            if (parentGroup) {
              parentGroup.children.push(svgItem);
            } else {
              items.push(svgItem);
            }
          } catch (e) {
            // пропускаем узлы, которые Figma не может экспортировать (нет видимых слоёв и т.п.)
          }
        }
        else if (hasImageFill(node)) {
          if (!hasRenderableSize(node)) return;
          try {
            const bytes = await (node as unknown as ExportMixin).exportAsync({ format: "PNG", constraint: { type: "SCALE", value: 1 }});
            const dataBase64 = "data:image/png;base64," + figma.base64Encode(bytes);
            items.push({
              id: node.id,
              type: "IMAGE",
              xIn, yIn, wIn, hIn,
              rotate,
              mime: "image/png",
              dataBase64,
              name: node.name,
              zIndex: zCounter++
            });
          } catch (e) {
            // пропускаем узлы, которые Figma не может экспортировать
          }
        }
      };
      await walk(frame as unknown as SceneNode);
  
      exportFrames.push({
        id: frame.id,
        name: frame.name,
        widthPx: frame.width,
        heightPx: frame.height,
        slideWIn, slideHIn,
        bgColorHex: bgHex || undefined,
        items
      });
    }
  
    figma.ui.onmessage = (msg) => {
      if (msg.type === "done") {
        figma.closePlugin("Export complete");
      }
    };

    // Без UI: отображение выключено, сразу отправляем данные для авто-экспорта
    figma.showUI(__html__, { visible: false });
    const fileNameBase = exportFrames.length === 1 ? exportFrames[0].name : figma.currentPage.name;
    figma.ui.postMessage({ type: "AUTO", frames: exportFrames, fileNameBase });
  });
  
  // ==== helpers ====
  
  async function traverse(node: SceneNode, fn: (n: SceneNode) => void | Promise<void>) : Promise<void> {
    // Сохраняем порядок слоёв: от заднего к переднему
    if ("children" in node) {
      for (const child of node.children as SceneNode[]) {
        await traverse(child, fn);
      }
    }
    // Пропускаем сам фрейм как геом. элемент, но обрабатываем содержимое
    if (node.type !== "FRAME" && node.type !== "COMPONENT" && node.type !== "INSTANCE") {
      await fn(node);
    }
  }
  
  function absOrigin(n: SceneNode) {
    // absoluteTransform = [[a,b,tx],[c,d,ty]]
    const m = (n as any).absoluteTransform as Transform;
    return { x: m[0][2], y: m[1][2] };
  }
  
  function rotationFromTransform(m: Transform): number {
    const a = m[0][0], b = m[0][1], c = m[1][0], d = m[1][1];
    
    // Вычисляем угол поворота из матрицы трансформации
    let angle = Math.atan2(b, a) * 180 / Math.PI;
    
    // Нормализуем угол в диапазон -180 до 180
    while (angle > 180) angle -= 360;
    while (angle < -180) angle += 360;
    
    console.log('Transform matrix:', m, 'Calculated angle:', angle);
    
    return Math.round(angle * 1000) / 1000;
  }
  
  function firstSolidPaint(paints: ReadonlyArray<Paint> | PluginAPI["mixed"]) {
    if (!Array.isArray(paints)) return null;
    const p = paints.find(p => p.type === "SOLID") as SolidPaint | undefined;
    return p || null;
  }

  function frameBackgroundHex(n: FrameNode | ComponentNode | InstanceNode) {
    const paints = (n as FrameNode).backgrounds as ReadonlyArray<Paint> | PluginAPI["mixed"];
    const p = firstSolidPaint(paints);
    if (!p) return null;
    return rgbToHex(p.color, (p.opacity === undefined || p.opacity === null) ? 1 : p.opacity);
  }
  
  function rgbToHex(c: RGB, opacity: number = 1) {
    const r = Math.round(c.r * 255);
    const g = Math.round(c.g * 255);
    const b = Math.round(c.b * 255);
    // Игнорируем прозрачность для текста — PowerPoint не поддерживает прозрачные глифы напрямую
    return [r, g, b].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  
  function mapAlign(a: "LEFT" | "CENTER" | "RIGHT" | "JUSTIFIED"): "left" | "center" | "right" {
    return a === "CENTER" ? "center" : a === "RIGHT" ? "right" : "left";
  }
  
  function mapVAlign(v: "TOP" | "CENTER" | "BOTTOM"): "top" | "middle" | "bottom" {
    return v === "CENTER" ? "middle" : v === "BOTTOM" ? "bottom" : "top";
  }
  
  function isFont(f: FontName | PluginAPI["mixed"]): f is FontName {
    return !!f && typeof f === "object" && "family" in f;
  }
  
  function isVectorNode(n: SceneNode): boolean {
    return (
      (n.type === "VECTOR" ||
        n.type === "LINE" ||
        n.type === "ELLIPSE" ||
        n.type === "POLYGON" ||
        n.type === "STAR" ||
        n.type === "RECTANGLE" ||
        n.type === "BOOLEAN_OPERATION" ||
        (n.type === "GROUP" && (n as any).vectorNetwork)) &&
      "exportAsync" in n
    );
  }

  function hasImageFill(n: SceneNode): boolean {
    // Считаем растровыми только узлы с заливкой типа IMAGE
    const anyNode = n as any;
    const fills = anyNode && anyNode.fills;
    if (!fills || !Array.isArray(fills)) return false;
    return fills.some((p: any) => p && p.type === 'IMAGE');
  }

  function hasSize(n: SceneNode): n is SceneNode & { width: number; height: number } {
    return "width" in n && "height" in n;
  }

  function hasRenderableSize(n: SceneNode): boolean {
    if (!("exportAsync" in n)) return false;
    if (!hasSize(n)) return false;
    // Разрешаем линии (w == 0 или h == 0), но отсекаем полностью нулевые
    if (n.width === 0 && n.height === 0) return false;
    // Контейнеры не рендерим как отдельные картинки
    if (n.type === "GROUP" || n.type === "FRAME" || n.type === "COMPONENT" || n.type === "INSTANCE") {
      const g = n as GroupNode;
      if (!('children' in g) || !g.children || g.children.length === 0) return false;
      return false;
    }
    return true;
  }
  
  function pxToIn(px: number):number { return px / 96; }       // 96 px = 1 in
  
  function pxToPt(px: number):number { return px * 0.75; }     // 1 px = 0.75 pt
  
  function frameBorder(n: FrameNode | ComponentNode | InstanceNode) {
    const strokes = (n as any).strokes as ReadonlyArray<Paint> | PluginAPI["mixed"];
    const strokeWeight = (n as any).strokeWeight;
    const strokeAlign = (n as any).strokeAlign;
    
    if (!strokes || !Array.isArray(strokes) || !strokeWeight) return null;
    
    const stroke = firstSolidPaint(strokes);
    if (!stroke) return null;
    
    return {
      color: rgbToHex(stroke.color, (stroke.opacity === undefined || stroke.opacity === null) ? 1 : stroke.opacity),
      width: pxToPt(strokeWeight),
      style: strokeAlign === 'CENTER' ? 'solid' : 'solid' // можно расширить для разных стилей
    };
  }