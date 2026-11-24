/**
 * docx 库在“没有 atob”的环境（即 Node）里才会走 require("buffer") 分支。
 * 浏览器里 atob 一定存在，那段代码是死代码，这里提供一个占位实现，
 * 让打包器可以正常解析这个 Node 内置模块。
 */
export class Buffer {
  constructor() {
    throw new Error("Buffer is not available in the browser build.");
  }
}

export default { Buffer };
