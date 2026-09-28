export const uploadFormData = (url, body, onProgress, Request = XMLHttpRequest) => new Promise((resolve, reject) => {
  const request = new Request();
  request.open('POST', url);
  request.responseType = 'text';
  request.upload.addEventListener('progress', event => {
    if (event.lengthComputable) onProgress?.(event.loaded, event.total);
  });
  request.addEventListener('load', () => {
    let data = {};
    try { data = request.responseText ? JSON.parse(request.responseText) : {}; }
    catch { reject(new Error('服务器返回了无法解析的响应')); return; }
    if (request.status >= 200 && request.status < 300) resolve(data);
    else reject(new Error(data.detail || request.statusText || '上传失败'));
  });
  request.addEventListener('error', () => reject(new Error('网络连接中断，上传未完成')));
  request.addEventListener('abort', () => reject(new DOMException('上传已取消', 'AbortError')));
  request.send(body);
});
