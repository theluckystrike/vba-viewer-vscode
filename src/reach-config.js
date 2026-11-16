'use strict';
// REACH1: this listing's Pro tools for the viewer toolbar row and the Explorer entries (src/reach.js).
module.exports = {
  prefix: 'vbaViewer',
  triesKey: 'vbaViewer.proTrial',
  slot: '#toolbar',
  tools: [
    { id: 'exportModules', title: 'Export all modules', command: 'vbaViewer.exportModules' },
    { id: 'scan', title: 'Macro security scan', command: 'vbaViewer.scan' },
    { id: 'compare', title: 'Compare macros', command: 'vbaViewer.compare' },
  ],
};
