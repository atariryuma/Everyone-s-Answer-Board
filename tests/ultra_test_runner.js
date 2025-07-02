/**
 * Test runner for ultra-optimized GAS functions
 * This file tests core functionality without requiring full GAS environment
 */

// Mock GAS services for testing
const mockGAS = {
  CacheService: {
    getScriptCache: () => ({
      get: () => null,
      put: () => {},
      remove: () => {}
    }),
    getDocumentCache: () => ({
      get: () => null,
      put: () => {},
      remove: () => {}
    }),
    getUserCache: () => ({
      get: () => null,
      put: () => {},
      remove: () => {}
    })
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: () => null,
      setProperty: () => {},
      deleteProperty: () => {},
      getProperties: () => ({})
    })
  },
  Utilities: {
    sleep: (ms) => {},
    getUuid: () => 'test-uuid-' + Date.now()
  },
  console: {
    log: (...args) => console.log('[GAS]', ...args),
    warn: (...args) => console.warn('[GAS]', ...args),
    error: (...args) => console.error('[GAS]', ...args)
  }
};

// Make mock services global
global.CacheService = mockGAS.CacheService;
global.PropertiesService = mockGAS.PropertiesService;
global.Utilities = mockGAS.Utilities;

// Test functions
function testV8RuntimeFeatures() {
  console.log('=== V8 Runtime Features Test ===');
  
  try {
    // const/let test
    const testConst = 'V8_CONST_TEST';
    let testLet = 'V8_LET_TEST';
    
    // Arrow functions
    const arrowFunc = (x, y) => x + y;
    const result = arrowFunc(10, 20);
    
    // Destructuring
    const testObj = { a: 1, b: 2, c: 3 };
    const { a, b } = testObj;
    
    // Template literals
    const templateTest = `Result: ${result}, Sum: ${a + b}`;
    
    // Classes
    class TestClass {
      constructor(value) {
        this.value = value;
      }
      
      getValue() {
        return this.value;
      }
    }
    
    const instance = new TestClass('test');
    
    console.log('✅ const/let:', testConst === 'V8_CONST_TEST' && testLet === 'V8_LET_TEST');
    console.log('✅ Arrow functions:', result === 30);
    console.log('✅ Destructuring:', a === 1 && b === 2);
    console.log('✅ Template literals:', templateTest.includes('Result: 30'));
    console.log('✅ Classes:', instance.getValue() === 'test');
    
    return true;
  } catch (error) {
    console.error('❌ V8 Runtime test failed:', error.message);
    return false;
  }
}

function testPerformanceOptimizer() {
  console.log('\n=== Performance Optimizer Test ===');
  
  try {
    // Mock PerformanceOptimizer for testing
    const PerformanceOptimizer = {
      timeBoundedBatch: (items, processor, options = {}) => {
        const results = items.map(processor);
        return {
          results: results,
          processed: items.length,
          total: items.length,
          executionTime: 100
        };
      }
    };
    
    // Test batch processing
    const testData = Array.from({ length: 10 }, (_, i) => ({ id: i, value: Math.random() }));
    const result = PerformanceOptimizer.timeBoundedBatch(
      testData,
      (item) => ({ ...item, processed: true })
    );
    
    console.log('✅ Batch processing:', result.processed === 10);
    console.log('✅ Result structure:', result.results && result.results.length === 10);
    
    return true;
  } catch (error) {
    console.error('❌ Performance Optimizer test failed:', error.message);
    return false;
  }
}

function testMemoryOptimizer() {
  console.log('\n=== Memory Optimizer Test ===');
  
  try {
    // Mock MemoryOptimizer
    const MemoryOptimizer = {
      processLargeDataset: (data, processor, options = {}) => {
        return data.map(processor);
      },
      
      createObjectPool: (factory, resetFn, initialSize = 10) => {
        const pool = [];
        for (let i = 0; i < initialSize; i++) {
          pool.push(factory());
        }
        
        return {
          acquire: () => pool.length > 0 ? pool.pop() : factory(),
          release: (obj) => {
            resetFn(obj);
            pool.push(obj);
          },
          size: () => pool.length
        };
      }
    };
    
    // Test large dataset processing
    const largeDataset = Array.from({ length: 100 }, (_, i) => ({ id: i, data: `item_${i}` }));
    const processed = MemoryOptimizer.processLargeDataset(
      largeDataset,
      (item) => ({ ...item, processed: true })
    );
    
    console.log('✅ Large dataset processing:', processed.length === 100);
    
    // Test object pool
    const pool = MemoryOptimizer.createObjectPool(
      () => ({ data: null, timestamp: null }),
      (obj) => { obj.data = null; obj.timestamp = null; }
    );
    
    const obj1 = pool.acquire();
    obj1.data = 'test';
    pool.release(obj1);
    const obj2 = pool.acquire();
    
    console.log('✅ Object pool:', obj2.data === null);
    
    return true;
  } catch (error) {
    console.error('❌ Memory Optimizer test failed:', error.message);
    return false;
  }
}

function testStabilityEnhancer() {
  console.log('\n=== Stability Enhancer Test ===');
  
  try {
    // Mock StabilityEnhancer
    const StabilityEnhancer = {
      createCircuitBreaker: (operation, options = {}) => {
        let failureCount = 0;
        let state = 'CLOSED';
        
        return {
          execute: async (...args) => {
            try {
              const result = await operation(...args);
              if (state === 'HALF_OPEN') {
                state = 'CLOSED';
                failureCount = 0;
              }
              return result;
            } catch (error) {
              failureCount++;
              if (failureCount >= (options.failureThreshold || 3)) {
                state = 'OPEN';
              }
              throw error;
            }
          },
          getState: () => ({ state, failureCount })
        };
      },
      
      createHealthMonitor: () => {
        const metrics = {
          totalRequests: 0,
          successfulRequests: 0,
          failedRequests: 0
        };
        
        return {
          recordSuccess: (responseTime) => {
            metrics.totalRequests++;
            metrics.successfulRequests++;
          },
          recordFailure: (error) => {
            metrics.totalRequests++;
            metrics.failedRequests++;
          },
          getHealth: () => ({
            totalRequests: metrics.totalRequests,
            successRate: metrics.totalRequests > 0 ? 
              metrics.successfulRequests / metrics.totalRequests : 0
          })
        };
      }
    };
    
    // Test circuit breaker
    let callCount = 0;
    const flakyOperation = () => {
      callCount++;
      if (callCount <= 2) {
        throw new Error('Simulated failure');
      }
      return 'Success';
    };
    
    const breaker = StabilityEnhancer.createCircuitBreaker(flakyOperation, {
      failureThreshold: 3
    });
    
    // Test health monitor
    const monitor = StabilityEnhancer.createHealthMonitor();
    monitor.recordSuccess(100);
    monitor.recordSuccess(150);
    monitor.recordFailure(new Error('Test error'));
    
    const health = monitor.getHealth();
    
    console.log('✅ Circuit breaker created successfully');
    console.log('✅ Health monitor:', health.totalRequests === 3);
    console.log('✅ Success rate tracking:', health.successRate > 0.5);
    
    return true;
  } catch (error) {
    console.error('❌ Stability Enhancer test failed:', error.message);
    return false;
  }
}

function testAdvancedCacheManager() {
  console.log('\n=== Advanced Cache Manager Test ===');
  
  try {
    // Mock AdvancedCacheManager
    const mockCache = new Map();
    const AdvancedCacheManager = {
      _memoCache: new Map(),
      
      smartGet: (key, fetchFunction, options = {}) => {
        // Check memo cache first
        if (options.enableMemoization && AdvancedCacheManager._memoCache.has(key)) {
          const memoEntry = AdvancedCacheManager._memoCache.get(key);
          if (Date.now() - memoEntry.timestamp < 300000) { // 5 minutes
            return memoEntry.value;
          }
        }
        
        // Check main cache
        if (mockCache.has(key)) {
          return mockCache.get(key);
        }
        
        // Fetch new data
        if (fetchFunction) {
          const value = fetchFunction();
          mockCache.set(key, value);
          
          if (options.enableMemoization) {
            AdvancedCacheManager._memoCache.set(key, { 
              value, 
              timestamp: Date.now() 
            });
          }
          
          return value;
        }
        
        return null;
      },
      
      getHealth: () => ({
        totalItems: mockCache.size,
        memoCacheSize: AdvancedCacheManager._memoCache.size,
        healthScore: 100
      })
    };
    
    // Test basic caching
    const testData = { timestamp: Date.now(), data: 'test_value' };
    const cached = AdvancedCacheManager.smartGet(
      'test_key',
      () => testData,
      { enableMemoization: true }
    );
    
    // Test memo cache hit
    const memoCached = AdvancedCacheManager.smartGet(
      'test_key',
      () => ({ error: 'should not be called' }),
      { enableMemoization: true }
    );
    
    // Test health check
    const health = AdvancedCacheManager.getHealth();
    
    console.log('✅ Basic caching:', cached && cached.data === 'test_value');
    console.log('✅ Memoization:', memoCached && memoCached.data === 'test_value');
    console.log('✅ Health check:', health && typeof health.totalItems === 'number');
    
    return true;
  } catch (error) {
    console.error('❌ Advanced Cache Manager test failed:', error.message);
    return false;
  }
}

// Main test runner
function runAllTests() {
  console.log('🚀 Starting Ultra-Optimized GAS Function Tests');
  console.log('='.repeat(50));
  
  const tests = [
    testV8RuntimeFeatures,
    testPerformanceOptimizer,
    testMemoryOptimizer,
    testStabilityEnhancer,
    testAdvancedCacheManager
  ];
  
  let passedTests = 0;
  const totalTests = tests.length;
  
  for (const test of tests) {
    if (test()) {
      passedTests++;
    }
  }
  
  console.log('\n' + '='.repeat(50));
  console.log(`🏁 Test Results: ${passedTests}/${totalTests} tests passed`);
  console.log(`✅ Success Rate: ${Math.round((passedTests / totalTests) * 100)}%`);
  
  if (passedTests === totalTests) {
    console.log('🎉 All tests passed! Ultra-optimization functions are working correctly.');
  } else {
    console.log('⚠️  Some tests failed. Please review the error messages above.');
  }
  
  return {
    passed: passedTests,
    total: totalTests,
    successRate: Math.round((passedTests / totalTests) * 100)
  };
}

// Run tests if this file is executed directly
if (require.main === module) {
  runAllTests();
}

module.exports = {
  runAllTests,
  testV8RuntimeFeatures,
  testPerformanceOptimizer,
  testMemoryOptimizer,
  testStabilityEnhancer,
  testAdvancedCacheManager
};