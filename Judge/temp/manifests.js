(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else if(typeof exports === 'object')
		exports["debugManifests"] = factory();
	else
		root["debugManifests"] = factory();
})(window, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Set the webpack public path
/******/ 	(function () {
/******/ 	  var scripts = document.getElementsByTagName('script');
/******/ 	  var regex = new RegExp('manifests\\.js', 'i');
/******/ 	  var publicPath;
/******/ 	
/******/ 	  if (scripts && scripts.length) {
/******/ 	    for (var i = 0; i < scripts.length; i++) {
/******/ 	      if (!scripts[i]) continue;
/******/ 	      var path = scripts[i].getAttribute('src');
/******/ 	      if (path && path.match(regex)) {
/******/ 	        publicPath = path.substring(0, path.lastIndexOf('/') + 1);
/******/ 	        break;
/******/ 	      }
/******/ 	    }
/******/ 	  }
/******/ 	
/******/ 	  if (!publicPath) {
/******/ 	    for (var global in window.__setWebpackPublicPathLoaderSrcRegistry__) {
/******/ 	      if (global && global.match(regex)) {
/******/ 	        publicPath = global.substring(0, global.lastIndexOf('/') + 1);
/******/ 	        break;
/******/ 	      }
/******/ 	    }
/******/ 	  }
/******/ 	  __webpack_require__.p = publicPath;
/******/ 	})();
/******/ 	
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getManifests", function() { return getManifests; });
var MANIFESTS_ARRAY = [
  {
    "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a",
    "alias": "SPLodashSubset",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-lodash-subset",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-lodash-subset/"
      ],
      "scriptResources": {
        "sp-lodash-subset": {
          "type": "path",
          "path": "dist/sp-lodash-subset.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b",
    "alias": "SPCoreLibrary",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-core-library",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-core-library/"
      ],
      "scriptResources": {
        "sp-core-library": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-core-library_en-us.js"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "29bd516f-4ece-40b7-8028-597cbc65a223",
    "alias": "SpOfficeUIFabricCore",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "office-ui-fabric-core",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-office-ui-fabric-core/"
      ],
      "scriptResources": {
        "office-ui-fabric-core": {
          "type": "path",
          "path": "dist/office-ui-fabric-core.js"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "f97266fb-ccb7-430e-9384-4124d05295d3",
    "alias": "Decorators",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "decorators",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/decorators/"
      ],
      "scriptResources": {
        "decorators": {
          "type": "path",
          "path": "dist/decorators.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "manifestVersion": 2,
    "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8",
    "alias": "SPDiagnostics",
    "componentType": "Library",
    "version": "1.9.1",
    "loaderConfig": {
      "entryModuleId": "sp-diagnostics",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-diagnostics/"
      ],
      "scriptResources": {
        "sp-diagnostics": {
          "type": "path",
          "path": "dist/sp-diagnostics.js"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        }
      }
    },
    "isInternal": true
  },
  {
    "manifestVersion": 2,
    "id": "e40f8203-b39d-425a-a957-714852e33b79",
    "alias": "SPDynamicData",
    "componentType": "Library",
    "version": "1.9.1",
    "loaderConfig": {
      "entryModuleId": "sp-dynamic-data",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-dynamic-data/"
      ],
      "scriptResources": {
        "sp-dynamic-data": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-dynamic-data_en-us.js"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "1e384972-6346-49b4-93c7-b2e6763938e6",
    "alias": "sp-polyfills",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-polyfills",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-polyfills/"
      ],
      "scriptResources": {
        "sp-polyfills": {
          "type": "path",
          "path": "dist/sp-polyfills.js"
        }
      }
    }
  },
  {
    "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6",
    "alias": "SPHttp",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "preloadComponents": [
      "4d5eb168-6729-49a8-aec7-0e397f486b6e"
    ],
    "loaderConfig": {
      "entryModuleId": "sp-http",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-http/"
      ],
      "scriptResources": {
        "sp-http": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-http_en-us.js"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "02a01e42-69ab-403d-8a16-acd128661f8e",
    "alias": "OfficeUIFabricReact",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "office-ui-fabric-react-bundle",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/office-ui-fabric-react-bundle/"
      ],
      "scriptResources": {
        "office-ui-fabric-react-bundle": {
          "type": "path",
          "path": "dist/office-ui-fabric-react-bundle.js"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8",
    "alias": "SPPageContext",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-page-context",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-page-context/"
      ],
      "scriptResources": {
        "sp-page-context": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-page-context_en-us.js"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-dynamic-data": {
          "type": "component",
          "version": "1.9.1",
          "id": "e40f8203-b39d-425a-a957-714852e33b79"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f",
    "alias": "SPLoader",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-loader",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-loader/"
      ],
      "scriptResources": {
        "sp-loader": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-loader_en-us.js"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-dynamic-data": {
          "type": "component",
          "version": "1.9.1",
          "id": "e40f8203-b39d-425a-a957-714852e33b79"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.9.1",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@ms/sp-telemetry": {
          "type": "component",
          "version": "0.6.9",
          "id": "8217e442-8ed3-41fd-957d-b112e841286a"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6",
    "alias": "SPComponentBase",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-component-base",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-component-base/"
      ],
      "scriptResources": {
        "sp-component-base": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-component-base_en-us.js"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.9.1",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-dynamic-data": {
          "type": "component",
          "version": "1.9.1",
          "id": "e40f8203-b39d-425a-a957-714852e33b79"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.9.1",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "f9e737b7-f0df-4597-ba8c-3060f82380db",
    "alias": "SPPropertyPane",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-property-pane",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-property-pane/"
      ],
      "scriptResources": {
        "sp-property-pane": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-property-pane_en-us.js"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.9.1",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/office-ui-fabric-react-bundle": {
          "type": "component",
          "version": "1.9.1",
          "id": "02a01e42-69ab-403d-8a16-acd128661f8e"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "974a7777-0990-4136-8fa6-95d80114c2e0",
    "alias": "SPWebPartBase",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "preloadComponents": [
      "f9e737b7-f0df-4597-ba8c-3060f82380db"
    ],
    "loaderConfig": {
      "entryModuleId": "sp-webpart-base",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-webpart-base/"
      ],
      "scriptResources": {
        "sp-webpart-base": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-webpart-base_en-us.js"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.9.1",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.9.1",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/sp-loader": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@ms/sp-load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@ms/sp-telemetry": {
          "type": "component",
          "version": "0.6.9",
          "id": "8217e442-8ed3-41fd-957d-b112e841286a"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "0773bd53-a69e-4293-87e6-ba80ea4d614b",
    "alias": "SPExtensionBase",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-extension-base",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-extension-base/"
      ],
      "scriptResources": {
        "sp-extension-base": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-extension-base_en-us.js"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.9.1",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-loader": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f"
        },
        "@ms/sp-telemetry": {
          "type": "component",
          "version": "0.6.9",
          "id": "8217e442-8ed3-41fd-957d-b112e841286a"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "4df9bb86-ab0a-4aab-ab5f-48bf167048fb",
    "alias": "SPApplicationBase",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "preloadComponents": [
      "c0c518b8-701b-4f6f-956d-5782772bb731"
    ],
    "loaderConfig": {
      "entryModuleId": "sp-application-base",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-application-base/"
      ],
      "scriptResources": {
        "sp-application-base": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-application-base_en-us.js"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.9.1",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-extension-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "0773bd53-a69e-4293-87e6-ba80ea4d614b"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.9.1",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/sp-loader": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@ms/sp-load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@ms/sp-telemetry": {
          "type": "component",
          "version": "0.6.9",
          "id": "8217e442-8ed3-41fd-957d-b112e841286a"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "4d5eb168-6729-49a8-aec7-0e397f486b6e",
    "alias": "SPClientPreview",
    "componentType": "Library",
    "version": "1.9.1",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-client-preview",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-client-preview/"
      ],
      "scriptResources": {
        "sp-client-preview": {
          "type": "path",
          "path": "dist/sp-client-preview.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "8be81a5c-af38-4bb2-af97-afa3b64dfbed",
    "alias": "WebPartWorkbench",
    "componentType": "Application",
    "version": "1.9.1",
    "manifestVersion": 2,
    "title": {
      "default": "WebpartWorkbench"
    },
    "description": {
      "default": "WebpartWorkbench"
    },
    "assemblyId": "5dae53c4-db1e-4d0b-b8b2-88c874dabf83",
    "preloadComponents": [],
    "preloadOptions": {
      "shouldPreloadWeb": true,
      "shouldPreloadUser": true,
      "shouldPreloadList": false,
      "shouldPreloadItem": true,
      "shouldPreloadQuickLaunch": true
    },
    "loaderConfig": {
      "entryModuleId": "sp-webpart-workbench",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-webpart-workbench/"
      ],
      "scriptResources": {
        "sp-webpart-workbench": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-webpart-workbench_en-us.js"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/office-ui-fabric-react-bundle": {
          "type": "component",
          "version": "1.9.1",
          "id": "02a01e42-69ab-403d-8a16-acd128661f8e"
        },
        "@microsoft/sp-application-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "4df9bb86-ab0a-4aab-ab5f-48bf167048fb"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.9.1",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.9.1",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/sp-loader": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@microsoft/sp-webpart-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "974a7777-0990-4136-8fa6-95d80114c2e0"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "b2ae421d-3292-4e52-92a6-69c76e71ed02",
    "alias": "JudgeVacationRequestWebPart",
    "componentType": "WebPart",
    "version": "0.0.1",
    "manifestVersion": 2,
    "requiresCustomScript": false,
    "supportedHosts": [
      "SharePointWebPart"
    ],
    "preconfiguredEntries": [
      {
        "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
        "group": {
          "default": "Other"
        },
        "title": {
          "default": "Judge-Vacation-Request"
        },
        "description": {
          "default": "Judge-Vacation-Request description"
        },
        "officeFabricIconFontName": "Page",
        "properties": {
          "description": "Judge-Vacation-Request"
        }
      }
    ],
    "loaderConfig": {
      "entryModuleId": "judge-vacation-request-web-part",
      "internalModuleBaseUrls": [
        "https://localhost:4321/"
      ],
      "scriptResources": {
        "judge-vacation-request-web-part": {
          "type": "path",
          "path": "dist/judge-vacation-request-web-part.js"
        },
        "JudgeVacationRequestWebPartStrings": {
          "defaultPath": "lib/webparts/judgeVacationRequest/loc/en-us.js",
          "type": "localizedPath"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.9.1",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.9.1",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/sp-loader": {
          "type": "component",
          "version": "1.9.1",
          "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.9.1",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-webpart-base": {
          "type": "component",
          "version": "1.9.1",
          "id": "974a7777-0990-4136-8fa6-95d80114c2e0"
        }
      }
    }
  }
];
/**
 * Get the manifest array.
 */
function getManifests() {
    // Clone manifestsArray
    var manifests = JSON.parse(JSON.stringify(MANIFESTS_ARRAY));
    var manifestsFileUrl = __webpack_require__.p;
    if (manifestsFileUrl && manifestsFileUrl !== '') {
        manifests.forEach(function (manifest) {
            if (!manifest.loaderConfig.internalModuleBaseUrls || manifest.loaderConfig.internalModuleBaseUrls.length === 0) {
                manifest.loaderConfig.internalModuleBaseUrls = [manifestsFileUrl];
            }
        });
    }
    else {
        console.error("Unable to determine " + "manifests.js" + " file URL. Using default base URL.");
    }
    return manifests;
}
//# sourceMappingURL=manifestsFile.js.map

/***/ })
/******/ ]);
});