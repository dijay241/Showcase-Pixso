// Типизация для Pixso Plugin API
// Основано на документации: https://pixso.cn/developer/en/

interface RGB {
  r: number;
  g: number;
  b: number;
}

interface RGBA extends RGB {
  a: number;
}

interface Paint {
  type: "SOLID" | "GRADIENT_LINEAR" | "GRADIENT_RADIAL" | "GRADIENT_ANGULAR" | "GRADIENT_DIAMOND" | "IMAGE" | "EMOJI";
  color?: RGB;
  opacity?: number;
  gradientStops?: Array<{
    color: RGB;
    position: number;
  }>;
  gradientTransform?: Transform;
}

interface Transform {
  [0]: [number, number, number];
  [1]: [number, number, number];
}

interface FontName {
  family: string;
  style: string;
}

interface LetterSpacing {
  value: number;
  unit: "PIXELS" | "PERCENT" | "PT";
}

interface LineHeight {
  value: number;
  unit: "PIXELS" | "PERCENT" | "PT" | "AUTO";
}

interface ExportSettings {
  format: "PNG" | "JPG" | "SVG" | "PDF";
  constraint?: {
    type: "SCALE" | "WIDTH" | "HEIGHT";
    value: number;
  };
  svgOutlineText?: boolean;
  svgIdAttribute?: boolean;
  svgSimplifyStroke?: boolean;
}

interface BaseNode {
  id: string;
  name: string;
  type: string;
  visible: boolean;
  removed: boolean;
  parent: (BaseNode & ChildrenMixin) | null;
  absoluteTransform: Transform;
  relativeTransform?: Transform;
  rotation?: number;
  
  // Plugin Data API
  getPluginData(key: string): string;
  setPluginData(key: string, value: string): void;
  getPluginDataKeys(): string[];
  getSharedPluginData(namespace: string, key: string): string;
  setSharedPluginData(namespace: string, key: string, value: string): void;
  getSharedPluginDataKeys(namespace: string): string[];
  
  // Relaunch Data API
  setRelaunchData(data: { [command: string]: string }): void;
  getRelaunchData(): { [command: string]: string };
  
  // Export API
  exportAsync(settings?: ExportSettings): Promise<Uint8Array>;
  exportJsonAsync(payload?: { withConnectLine?: boolean }): Promise<string>;
  exportHexAsync(payload?: { withConnectLine?: boolean }): Promise<string>;
  
  // Common methods
  clone(): BaseNode;
  remove(): void;
  toString(): string;
}

interface ChildrenMixin {
  children: BaseNode[];
}

interface GeometryMixin {
  width: number;
  height: number;
  fills: Paint[];
  strokes: Paint[];
  strokeWeight: number;
  strokeAlign: "INSIDE" | "OUTSIDE" | "CENTER";
  cornerRadius: number;
  topLeftRadius: number;
  topRightRadius: number;
  bottomLeftRadius: number;
  bottomRightRadius: number;
}

interface TextNode extends BaseNode, GeometryMixin {
  type: "TEXT";
  characters: string;
  textAlignHorizontal: "LEFT" | "CENTER" | "RIGHT" | "JUSTIFIED";
  textAlignVertical: "TOP" | "CENTER" | "BOTTOM";
  
  // Text range methods
  getRangeFontName(start: number, end: number): FontName;
  getRangeFontSize(start: number, end: number): number;
  getRangeTextDecoration(start: number, end: number): "NONE" | "UNDERLINE" | "STRIKETHROUGH";
  getRangeLetterSpacing(start: number, end: number): LetterSpacing;
  getRangeLineHeight(start: number, end: number): LineHeight;
}

interface FrameNode extends BaseNode, GeometryMixin, ChildrenMixin {
  type: "FRAME" | "COMPONENT" | "INSTANCE";
  layoutMode?: "NONE" | "HORIZONTAL" | "VERTICAL";
  primaryAxisSizingMode?: "FIXED" | "AUTO";
  counterAxisSizingMode?: "FIXED" | "AUTO";
  paddingLeft?: number;
  paddingRight?: number;
  paddingTop?: number;
  paddingBottom?: number;
  itemSpacing?: number;
}

interface RectangleNode extends BaseNode, GeometryMixin {
  type: "RECTANGLE";
}

interface EllipseNode extends BaseNode, GeometryMixin {
  type: "ELLIPSE";
}

interface VectorNode extends BaseNode, GeometryMixin {
  type: "VECTOR" | "POLYGON" | "STAR" | "BOOLEAN_OPERATION";
}

interface GroupNode extends BaseNode, ChildrenMixin {
  type: "GROUP";
}

interface LineNode extends BaseNode {
  type: "LINE";
  width: number;
  rotation: number;
}

interface PageNode extends BaseNode, ChildrenMixin {
  type: "PAGE";
  selection: BaseNode[];
}

interface DocumentNode extends BaseNode, ChildrenMixin {
  type: "DOCUMENT";
}

interface PixsoUI {
  onmessage: ((message: any) => void) | null;
  postMessage(message: any): void;
}

interface LoadingMessage {
  cancel(): void;
}

interface PixsoAPI {
  // Event handling
  on(event: "run", callback: (data: { command: string }) => void): void;
  
  // Document and page access
  currentPage: PageNode;
  root: DocumentNode;
  
  // UI management
  showUI(html: string, options?: { visible?: boolean; width?: number; height?: number }): void;
  ui: PixsoUI;
  
  // Notifications and feedback
  notify(message: string, options?: { timeout?: number; error?: boolean }): void;
  closePlugin(message?: string): void;
  loading(message: string): LoadingMessage;
  
  // Utility functions
  base64Encode(data: Uint8Array): string;
  base64Decode(data: string): Uint8Array;
  
  // Node operations
  flatten(nodes: BaseNode[]): BaseNode;
  
  // Viewport and selection
  viewport: {
    center: { x: number; y: number };
    zoom: number;
  };
  
  // Plugin data (global)
  getPluginData(key: string): string;
  setPluginData(key: string, value: string): void;
  getPluginDataKeys(): string[];
  getSharedPluginData(namespace: string, key: string): string;
  setSharedPluginData(namespace: string, key: string, value: string): void;
  getSharedPluginDataKeys(namespace: string): string[];
}

// Глобальное объявление
declare const pixso: PixsoAPI;

// Экспорт типов для использования в других файлах
export {
  PixsoAPI,
  BaseNode,
  TextNode,
  FrameNode,
  RectangleNode,
  EllipseNode,
  VectorNode,
  GroupNode,
  LineNode,
  PageNode,
  DocumentNode,
  Paint,
  Transform,
  FontName,
  LetterSpacing,
  LineHeight,
  ExportSettings,
  RGB,
  RGBA
};
