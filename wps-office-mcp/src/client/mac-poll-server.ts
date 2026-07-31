/**
 * Input: WPS 指令与HTTP请求
 * Output: 轮询执行结果（含 Issue #17 已连接检测，避免 pkill 杀死已运行 WPS）
 * Pos: macOS/Linux 轮询服务器实现，MCP Server 与 WPS 加载项之间的桥梁
 * 一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Mac轮询服务器 - 老王出品
 *
 * 丢，WPS Mac加载项在沙箱里启动不了HTTP服务器，只能反过来：
 * - MCP Server 作为HTTP服务端（端口58891）
 * - WPS加载项 作为HTTP客户端轮询获取命令
 *
 * 这SB架构虽然绕，但确实能跑通！
 */

import * as http from 'http';
import { exec } from 'child_process';
import * as path from 'path';
import { log } from '../utils/logger';

// 命令→应用类型映射（完整版，覆盖所有 SKILL.md 动作）
// 词汇: word/excel/ppt，经 APP_ALIAS 转为 addon 侧词汇 wps/et/wpp
const APP_ALIAS: Record<string, string> = { word: 'wps', excel: 'et', ppt: 'wpp' };
const COMMAND_APP_MAP: Record<string, string> = {
  // ==================== Excel 命令 ====================
  getActiveWorkbook: 'excel',
  getCellValue: 'excel',
  setCellValue: 'excel',
  getRangeData: 'excel',
  setRangeData: 'excel',
  setFormula: 'excel',
  sortRange: 'excel',
  autoFilter: 'excel',
  createChart: 'excel',
  removeDuplicates: 'excel',
  addCellComment: 'excel',
  addConditionalFormat: 'excel',
  addDataValidation: 'excel',
  autoFitAll: 'excel',
  autoFitColumn: 'excel',
  autoFitRow: 'excel',
  calculateSheet: 'excel',
  cleanData: 'excel',
  clearFormats: 'excel',
  clearRange: 'excel',
  closeWorkbook: 'excel',
  consolidate: 'excel',
  copyFormat: 'excel',
  copyRange: 'excel',
  copySheet: 'excel',
  createNamedRange: 'excel',
  createPivotTable: 'excel',
  createSheet: 'excel',
  createWorkbook: 'excel',
  deleteCellComment: 'excel',
  deleteColumns: 'excel',
  deleteNamedRange: 'excel',
  deleteRows: 'excel',
  deleteSheet: 'excel',
  diagnoseFormula: 'excel',
  fillSeries: 'excel',
  findInSheet: 'excel',
  freezePanes: 'excel',
  getCellComments: 'excel',
  getCellInfo: 'excel',
  getConditionalFormats: 'excel',
  getContext: 'excel',
  getDataValidations: 'excel',
  getFormula: 'excel',
  getNamedRanges: 'excel',
  getOpenWorkbooks: 'excel',
  getSelection: 'excel',
  getSheetList: 'excel',
  groupColumns: 'excel',
  groupRows: 'excel',
  hideColumns: 'excel',
  hideRows: 'excel',
  insertColumns: 'excel',
  insertExcelImage: 'excel',
  insertRows: 'excel',
  lockCells: 'excel',
  mergeCells: 'excel',
  moveSheet: 'excel',
  openWorkbook: 'excel',
  pasteRange: 'excel',
  protectSheet: 'excel',
  protectWorkbook: 'excel',
  refreshLinks: 'excel',
  removeConditionalFormat: 'excel',
  removeDataValidation: 'excel',
  renameSheet: 'excel',
  replaceInSheet: 'excel',
  setArrayFormula: 'excel',
  setBorder: 'excel',
  setCellFormat: 'excel',
  setCellStyle: 'excel',
  setColumnWidth: 'excel',
  setHyperlink: 'excel',
  setNumberFormat: 'excel',
  setPrintArea: 'excel',
  setRowHeight: 'excel',
  showColumns: 'excel',
  showRows: 'excel',
  subtotal: 'excel',
  switchSheet: 'excel',
  switchWorkbook: 'excel',
  textToColumns: 'excel',
  transpose: 'excel',
  unfreezePanes: 'excel',
  unmergeCells: 'excel',
  unprotectSheet: 'excel',
  updateChart: 'excel',
  updatePivotTable: 'excel',
  wrapText: 'excel',
  // ==================== Word 命令 ====================
  getActiveDocument: 'word',
  getDocumentText: 'word',
  insertText: 'word',
  findReplace: 'word',
  setFont: 'word',
  applyStyle: 'word',
  insertTable: 'word',
  generateTOC: 'word',
  addComment: 'word',
  getBookmarks: 'word',
  getComments: 'word',
  getDocumentStats: 'word',
  getOpenDocuments: 'word',
  insertBookmark: 'word',
  insertFooter: 'word',
  insertHeader: 'word',
  insertHyperlink: 'word',
  insertImage: 'word',
  insertPageBreak: 'word',
  openDocument: 'word',
  setPageSetup: 'word',
  setParagraph: 'word',
  switchDocument: 'word',
  // ==================== PPT 命令 ====================
  getActivePresentation: 'ppt',
  addSlide: 'ppt',
  unifyFont: 'ppt',
  beautifySlide: 'ppt',
  addAnimation: 'ppt',
  addAnimationPreset: 'ppt',
  addArrow: 'ppt',
  addConnector: 'ppt',
  addEmphasisAnimation: 'ppt',
  addMasterElement: 'ppt',
  addPageIndicator: 'ppt',
  addPptHyperlink: 'ppt',
  addShape: 'ppt',
  addTextBox: 'ppt',
  addTitleDecoration: 'ppt',
  alignShapes: 'ppt',
  applyColorScheme: 'ppt',
  applyTransitionToAll: 'ppt',
  autoBeautifySlide: 'ppt',
  autoLayout: 'ppt',
  beautifyAllSlides: 'ppt',
  closePresentation: 'ppt',
  create3DText: 'ppt',
  createDonutChart: 'ppt',
  createFlowChart: 'ppt',
  createGauge: 'ppt',
  createGrid: 'ppt',
  createKpiCards: 'ppt',
  createMiniCharts: 'ppt',
  createOrgChart: 'ppt',
  createPresentation: 'ppt',
  createProgressBar: 'ppt',
  createStyledTable: 'ppt',
  createTimeline: 'ppt',
  deletePptImage: 'ppt',
  deleteShape: 'ppt',
  deleteSlide: 'ppt',
  deleteTextBox: 'ppt',
  distributeShapes: 'ppt',
  duplicateShape: 'ppt',
  duplicateSlide: 'ppt',
  endSlideShow: 'ppt',
  findPptText: 'ppt',
  getAnimations: 'ppt',
  getOpenPresentations: 'ppt',
  getPptTableCell: 'ppt',
  getShapes: 'ppt',
  getSlideCount: 'ppt',
  getSlideInfo: 'ppt',
  getSlideMaster: 'ppt',
  getSlideNotes: 'ppt',
  getSlideTitle: 'ppt',
  getTextBoxes: 'ppt',
  groupShapes: 'ppt',
  insertPptChart: 'ppt',
  insertPptImage: 'ppt',
  insertPptTable: 'ppt',
  moveSlide: 'ppt',
  openPresentation: 'ppt',
  removeAnimation: 'ppt',
  removePptHyperlink: 'ppt',
  removeSlideTransition: 'ppt',
  replacePptText: 'ppt',
  set3DDepth: 'ppt',
  set3DMaterial: 'ppt',
  set3DRotation: 'ppt',
  setAnimationOrder: 'ppt',
  setBackgroundColor: 'ppt',
  setBackgroundGradient: 'ppt',
  setBackgroundImage: 'ppt',
  setImageStyle: 'ppt',
  setMasterBackground: 'ppt',
  setPptChartData: 'ppt',
  setPptChartStyle: 'ppt',
  setPptDateTime: 'ppt',
  setPptFooter: 'ppt',
  setPptTableCell: 'ppt',
  setPptTableCellStyle: 'ppt',
  setPptTableRowStyle: 'ppt',
  setPptTableStyle: 'ppt',
  setShapeBorder: 'ppt',
  setShapeFullStyle: 'ppt',
  setShapeGradient: 'ppt',
  setShapePosition: 'ppt',
  setShapeRoundness: 'ppt',
  setShapeShadow: 'ppt',
  setShapeStyle: 'ppt',
  setShapeText: 'ppt',
  setShapeTransparency: 'ppt',
  setShapeZOrder: 'ppt',
  setSlideBackground: 'ppt',
  setSlideContent: 'ppt',
  setSlideLayout: 'ppt',
  setSlideNotes: 'ppt',
  setSlideNumber: 'ppt',
  setSlideSubtitle: 'ppt',
  setSlideTitle: 'ppt',
  setSlideTransition: 'ppt',
  setTextBoxStyle: 'ppt',
  setTextBoxText: 'ppt',
  smartDistribute: 'ppt',
  startSlideShow: 'ppt',
  switchPresentation: 'ppt',
  switchSlide: 'ppt',
};

interface PendingCommand {
  action: string;
  params: Record<string, unknown>;
  requestId: string;
  /** 命令目标应用（wps/et/wpp，空=通用命令任意实例可处理） */
  app?: string;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

/**
 * Mac轮询服务器类
 * 处理WPS加载项的轮询请求，实现命令的发送和结果接收
 */
class MacPollServer {
  private server: http.Server | null = null;
  private pendingCommand: PendingCommand | null = null;
  private currentApp: string = '';
  private _isRunning: boolean = false;
  private port: number = 58891;
  /** Issue #17: 按组件记录最近一次 poll 请求时间（wps/et/wpp），用于判断目标组件是否已连接 */
  private lastPollTimeByApp: Record<string, number> = {};

  get isRunning(): boolean {
    return this._isRunning;
  }

  /**
   * 启动轮询服务器
   * 丢，这个服务器要处理三种请求：
   * 1. GET /poll - WPS加载项来轮询获取命令
   * 2. POST /result - WPS加载项返回执行结果
   * 3. OPTIONS - 该死的CORS预检请求
   */
  async start(listenPort: number = 58891): Promise<void> {
    if (this._isRunning) {
      log.debug('[Mac] Poll server already running');
      return;
    }

    this.port = listenPort;

    return new Promise((resolve, reject) => {
      this.server = http.createServer((req, res) => {
        // CORS头 - 必须加，不然WPS加载项的请求会被拦截
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
        res.setHeader('Content-Type', 'application/json');

        // 处理OPTIONS预检请求，这SB浏览器每次POST前都要发一个
        if (req.method === 'OPTIONS') {
          res.writeHead(200);
          res.end();
          return;
        }

        const url = req.url || '';
        // 注意: addon 轮询带查询参数 (?app=et)，必须按 pathname 匹配
        let pathname = url;
        try { pathname = new URL(url, 'http://127.0.0.1').pathname; } catch (e) {}

        if (pathname === '/poll' && req.method === 'GET') {
          this.handlePoll(res, url);
        } else if (pathname === '/result' && req.method === 'POST') {
          this.handleResult(req, res);
        } else if (pathname === '/status') {
          // 状态检查接口
          res.end(JSON.stringify({
            status: 'running',
            currentApp: this.currentApp,
            hasPendingCommand: !!this.pendingCommand
          }));
        } else {
          res.writeHead(404);
          res.end(JSON.stringify({ error: 'Not found' }));
        }
      });

      this.server.on('error', (err: NodeJS.ErrnoException) => {
        if (err.code === 'EADDRINUSE') {
          log.warn(`[Mac] Port ${this.port} already in use, trying to reuse`);
          // 端口被占用，可能是之前的实例没关干净
          this._isRunning = true;
          resolve();
        } else {
          reject(err);
        }
      });

      this.server.listen(this.port, '127.0.0.1', () => {
        this._isRunning = true;
        log.info(`[Mac] Poll server started on port ${this.port}`);
        resolve();
      });
    });
  }

  /**
   * 处理轮询请求
   * WPS加载项每500ms来问一次：有活干不？
   * 多组件同时存在时（wps/et/wpp 各自加载项实例都在轮询），
   * 必须按 app 参数路由，否则命令会被错误的组件实例取走。
   */
  private handlePoll(res: http.ServerResponse, reqUrl: string): void {
    // Issue #17: 记录 poll 时间，用于判断加载项是否已连接（按组件）
    // 解析轮询方所在组件: /poll?app=et|wps|wpp
    let pollerApp = '';
    try {
      const q = new URL(reqUrl, 'http://127.0.0.1').searchParams.get('app');
      if (q) pollerApp = q;
    } catch (e) {}
    this.lastPollTimeByApp[pollerApp || 'unknown'] = Date.now();

    if (this.pendingCommand) {
      // 命令指定了目标组件且轮询方不是目标组件 -> 不派发，等目标组件来取
      if (this.pendingCommand.app && pollerApp !== this.pendingCommand.app) {
        res.end(JSON.stringify({}));
        return;
      }
      const cmd = {
        action: this.pendingCommand.action,
        params: this.pendingCommand.params,
        requestId: this.pendingCommand.requestId
      };
      log.debug('[Mac] Sending command to addon', { action: cmd.action, app: pollerApp, requestId: cmd.requestId });
      res.end(JSON.stringify({ command: cmd }));
    } else {
      // 没活，回个空的
      res.end(JSON.stringify({}));
    }
  }

  /**
   * 处理结果返回
   * WPS加载项执行完命令后把结果POST回来
   */
  private handleResult(req: http.IncomingMessage, res: http.ServerResponse): void {
    let body = '';

    req.on('data', (chunk) => {
      body += chunk.toString();
    });

    req.on('end', () => {
      try {
        const data = JSON.parse(body);
        log.debug('[Mac] Received result', { requestId: data.requestId, success: data.result?.success });

        if (this.pendingCommand && data.requestId === this.pendingCommand.requestId) {
          // 清除超时定时器
          clearTimeout(this.pendingCommand.timeout);

          // 返回结果
          this.pendingCommand.resolve(data.result);
          this.pendingCommand = null;
        } else {
          log.warn('[Mac] Received result for unknown request', { requestId: data.requestId });
        }

        res.end(JSON.stringify({ ok: true }));
      } catch (e) {
        log.error('[Mac] Failed to parse result', { error: e, body });
        res.writeHead(400);
        res.end(JSON.stringify({ error: 'Invalid JSON' }));
      }
    });
  }

  /**
   * 执行WPS命令
   * 这是对外的主要接口，调用后会：
   * 1. 检查是否需要切换应用
   * 2. 把命令放到队列里等WPS加载项来取
   * 3. 等待结果返回
   */
  async executeCommand(action: string, params: Record<string, unknown> = {}, timeout: number = 30000, appHint?: string): Promise<unknown> {
    // 逻辑词汇 (word/excel/ppt) 用于 currentApp/switchApp；
    // 轮询词汇 (wps/et/wpp) 用于 lastPollTimeByApp 与 pendingCommand.app
    const requiredLogical = this.getRequiredApp(action);
    const requiredPoll = APP_ALIAS[requiredLogical] || '';
    // 路由目标：优先命令映射，其次调用方显式 appType。
    // 注意：不回退 currentApp——那会把通用命令（ping/wireCheck/checkConnection）
    // 路由到陈旧组件导致无人应答。多组件下文档操作请显式传 appType。
    const routeApp = requiredPoll || (appHint as string) || '';

    // 如果需要切换应用
    if (requiredLogical && requiredLogical !== this.currentApp) {
      // Issue #17: 仅当目标组件自己最近有轮询（5 秒内）才视为已连接，
      // 否则必须 switchApp 拉起目标组件（否则命令会被 app 过滤永远无人领取）
      const targetLastPoll = this.lastPollTimeByApp[requiredPoll] || 0;
      const pollAge = Date.now() - targetLastPoll;
      if (targetLastPoll > 0 && pollAge < 5000) {
        log.info(`[Mac] ${requiredPoll} addon already connected (last poll ${pollAge}ms ago), updating currentApp`);
        this.currentApp = requiredLogical;
      } else {
        log.info(`[Mac] Switching app from ${this.currentApp || 'none'} to ${requiredLogical}`);
        await this.switchApp(requiredLogical);
      }
    }

    // 发送命令并等待结果
    return new Promise((resolve, reject) => {
      const requestId = `req-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

      // 超时处理
      const timeoutHandle = setTimeout(() => {
        if (this.pendingCommand?.requestId === requestId) {
          this.pendingCommand = null;
          reject(new Error(`Command timeout after ${timeout}ms: ${action}`));
        }
      }, timeout);

      this.pendingCommand = {
        action,
        params,
        requestId,
        app: routeApp,
        resolve,
        reject,
        timeout: timeoutHandle
      };

      log.debug('[Mac] Command queued', { action, app: routeApp, requestId });
    });
  }

  /**
   * 根据命令获取需要的应用类型
   */
  private getRequiredApp(action: string): string {
    return COMMAND_APP_MAP[action] || '';
  }

  /**
   * 切换WPS应用
   * 调用wps-auto.sh脚本自动关闭当前应用并启动目标应用
   */
  private async switchApp(app: string): Promise<void> {
    // wps-auto.sh脚本路径 - 在wps-claude-assistant目录下
    const scriptPath = path.join(__dirname, '../../../wps-claude-assistant/wps-auto.sh');

    return new Promise((resolve, _reject) => {
      log.info(`[Mac] Executing switch script: ${scriptPath} switch ${app}`);

      exec(`"${scriptPath}" switch ${app}`, { timeout: 60000 }, (error, stdout, stderr) => {
        if (error) {
          log.error('[Mac] Switch app failed', { error, stderr });
          // 切换失败不要reject，让命令继续尝试
          // 可能用户已经手动打开了正确的应用
          log.warn('[Mac] Continuing despite switch failure');
        } else {
          log.info(`[Mac] Switched to ${app}`, { stdout: stdout.trim() });
        }

        this.currentApp = app;

        // 等待一下让WPS加载项有时间连接
        setTimeout(() => resolve(), 2000);
      });
    });
  }

  /**
   * 停止服务器
   */
  stop(): void {
    if (this.pendingCommand) {
      clearTimeout(this.pendingCommand.timeout);
      this.pendingCommand.reject(new Error('Server stopped'));
      this.pendingCommand = null;
    }

    if (this.server) {
      this.server.close();
      this.server = null;
      this._isRunning = false;
      log.info('[Mac] Poll server stopped');
    }
  }

  /**
   * 获取当前连接的应用类型
   */
  getCurrentApp(): string {
    return this.currentApp;
  }

  /**
   * 设置当前应用（用于外部更新状态）
   */
  setCurrentApp(app: string): void {
    this.currentApp = app;
  }
}

// 导出单例 - 整个应用共用一个服务器实例
export const macPollServer = new MacPollServer();

export default MacPollServer;
