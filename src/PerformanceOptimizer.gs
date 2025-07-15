/**
 * PerformanceOptimizer.gs - 軽量版パフォーマンス監視
 */

var globalProfiler = {
  timers: {},
  start: function(name) {
    this.timers[name] = new Date().getTime();
  },
  end: function(name) {
    if (this.timers[name]) {
      var elapsed = new Date().getTime() - this.timers[name];
      console.log(`Profile ${name}: ${elapsed}ms`);
      delete this.timers[name];
    }
  }
};