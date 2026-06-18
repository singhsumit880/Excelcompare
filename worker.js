// worker.js - Background Web Worker Wrapper
self.importScripts('lib/xlsx.full.min.js');
self.importScripts('compare-engine.js?v=2');

self.onmessage = function (e) {
  const { fileABuffer, fileBBuffer, config } = e.data;
  
  try {
    const startTime = performance.now();
    
    const result = compareWorkbooks(fileABuffer, fileBBuffer, config, (message) => {
      postMessage({ status: 'progress', message: message });
    });
    
    const endTime = performance.now();
    const elapsedTime = ((endTime - startTime) / 1000).toFixed(2);
    result.summary.elapsedTime = `${elapsedTime} sec`;
    
    postMessage({
      status: 'complete',
      result: result
    });
    
  } catch (error) {
    postMessage({
      status: 'error',
      message: 'Comparison failed: ' + error.message
    });
  }
};
