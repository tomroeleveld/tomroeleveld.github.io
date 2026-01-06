var analyticsGa = (function (exports) {
  'use strict';

  function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object. getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

  function _objectSpread(target) { for (var i = 1; i < arguments. length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { _defineProperty(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ?  Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

  function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable:  true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

  /* SharePoint User Info Functions */
  var cachedUserInfo = null;

  function getCurrentUserInfo() {
    if (cachedUserInfo) {
      return cachedUserInfo;
    }

    var xhr = new XMLHttpRequest();
    xhr.open("GET", "/_api/Web/CurrentUser? $select=Email,Title,Id", false);
    xhr.setRequestHeader("Accept", "application/json;odata=verbose");
    
    try {
      xhr.send();
      
      if (xhr.status === 200) {
        var data = JSON.parse(xhr.responseText);
        cachedUserInfo = {
          email: data. d.Email,
          name: data.d.Title,
          id: data. d.Id
        };
        return cachedUserInfo;
      } else {
        console.error("Error fetching user info:", xhr.status);
        return null;
      }
    } catch (error) {
      console.error("Error in getCurrentUserInfo:", error);
      return null;
    }
  }

  function stringToArrayBuffer(str) {
    var buf = new ArrayBuffer(str. length * 2);
    var bufView = new Uint16Array(buf);
    for (var i = 0, strLen = str.length; i < strLen; i++) {
      bufView[i] = str.charCodeAt(i);
    }
    return buf;
  }

  function arrayBufferToHex(buffer) {
    var byteArray = new Uint8Array(buffer);
    var hexString = '';
    for (var i = 0; i < byteArray. length; i++) {
      var hex = byteArray[i]. toString(16);
      hexString += (hex.length === 1 ? '0' : '') + hex;
    }
    return hexString;
  }

  function getHashedUserId() {
    var userInfo = getCurrentUserInfo();
    
    if (!userInfo || !userInfo.email || !userInfo.id) {
      console.error("Unable to get user info for hashing");
      return null;
    }

    try {
      // Combineer email met ID als salt
      var saltedEmail = userInfo.email + userInfo.id. toString();
      
      // Converteer naar ArrayBuffer
      var encoder = new TextEncoder();
      var data = encoder.encode(saltedEmail);
      
      // Maak synchrone hash (synchrone crypto is beperkt, dus we gebruiken een workaround)
      // Note: SubtleCrypto. digest is alleen async, dus we moeten een sync alternatief gebruiken
      // Voor een pure sync oplossing gebruiken we een eenvoudigere hash functie
      
      // Simpele hash functie (niet crypto-safe maar wel deterministisch)
      var hash = 0;
      for (var i = 0; i < saltedEmail.length; i++) {
        var char = saltedEmail.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // Convert to 32bit integer
      }
      
      // Converteer naar hex string
      var hashHex = Math.abs(hash).toString(16);
      
      // Pad met nullen voor consistente lengte
      while (hashHex.length < 8) {
        hashHex = '0' + hashHex;
      }
      
      return 'sp_' + hashHex + '_' + userInfo.id;
      
    } catch (error) {
      console.error("Error creating hashed user ID:", error);
      return null;
    }
  }

  /* global, window */
  var loadedInstances = {};
  /* Location of gtag script */

  var gtagScriptSource = 'https://www.googletagmanager.com/gtag/js'; // See https://developers.google.com/analytics/devguides/collection/ga4/reference/config

  var defaultGtagConf = {
    // https://support.google.com/analytics/answer/7201382? hl=en#zippy=%2Cglobal-site-tag-websites
    debug_mode: false,

    /**
     * Disable automatic sending of page views, instead let analytics. page() do this
     * https://developers.google.com/analytics/devguides/collection/gtagjs
     */
    send_page_view: false,
    // https://developers.google.com/analytics/devguides/collection/gtagjs/ip-anonymization
    anonymize_ip: false,

    /**
     * Disable All Advertising
     * https://developers.google.com/analytics/devguides/collection/ga4/display-features#disable_all_advertising_features
     */
    allow_google_signals: true,

    /**
     * Disable Advertising Personalization
     * https://developers.google.com/analytics/devguides/collection/ga4/display-features#disable_advertising_personalization
     */
    allow_ad_personalization_signals: true,

    /**
     * https://developers.google.com/analytics/devguides/collection/gtagjs/cookies-user-id#configure_cookie_field_settings
     */
    // cookie_domain: 'auto',
    // cookie_expires
    // cookie_prefix
    // cookie_update
    // cookie_flags

    /**
     * Cookie Flags
     * https://developers.google.com/analytics/devguides/collection/ga4/cookies-user-id#cookie_flags
     */
    cookie_flags: ''
  };
  var defaultConfig = {
    gtagName: 'gtag',
    dataLayerName: 'ga4DataLayer',
    measurementIds: [],
    gtagConfig: defaultGtagConf
  };
  /**
   * Google analytics plugin
   * @link https://getanalytics.io/plugins/google-analytics/
   * @link https://analytics.google.com/analytics/web/
   * @link https://developers.google.com/analytics/devguides/collection/analyticsjs
   * @param {object}  pluginConfig - Plugin settings
   * @param {string[]} pluginConfig.measurementIds - Google Analytics MEASUREMENT IDs
   * @param {boolean} [pluginConfig.debug] - Enable Google Analytics debug mode
   * @param {string}  [pluginConfig.dataLayerName=ga4DataLayer] - The optional name for dataLayer object.  Defaults to ga4DataLayer. 
   * @param {string}  [pluginConfig.gtagName=gtag] - The optional name for dataLayer object. Defaults to `gtag`.
   * @param {boolean} [pluginConfig.gtagConfig.anonymize_ip] - Enable [Anonymizing IP addresses](https://bit.ly/3c660Rd) sent to Google Analytics. 
   * @param {object}  [pluginConfig.gtagConfig.cookie_domain] - Additional cookie properties for configuring the [ga cookie](https://developers.google.com/analytics/devguides/collection/analyticsjs/cookies-user-id#configuring_cookie_field_settings)
   * @param {object}  [pluginConfig.gtagConfig.cookie_expires] - Additional cookie properties for configuring the [ga cookie](https://developers.google.com/analytics/devguides/collection/analyticsjs/cookies-user-id#configuring_cookie_field_settings)
   * @param {object}  [pluginConfig.gtagConfig.cookie_prefix] - Additional cookie properties for configuring the [ga cookie](https://developers.google.com/analytics/devguides/collection/analyticsjs/cookies-user-id#configuring_cookie_field_settings)
   * @param {object}  [pluginConfig. gtagConfig.cookie_update] - Additional cookie properties for configuring the [ga cookie](https://developers.google.com/analytics/devguides/collection/analyticsjs/cookies-user-id#configuring_cookie_field_settings)
   * @param {object}  [pluginConfig.gtagConfig.cookie_flags] - Additional cookie properties for configuring the [ga cookie](https://developers.google.com/analytics/devguides/collection/analyticsjs/cookies-user-id#configuring_cookie_field_settings)
   * @param {string}  [pluginConfig.customScriptSrc] - Custom URL for google analytics script, if proxying calls
   * @param {string}  [pluginConfig.nonce] - Content-Security-Policy nonce value
   * @return {*}
   * @example
   *
   * googleAnalytics({
   *   measurementIds:  ['G-abc123']
   * })
   */

  function googleAnalytics() {
    var pluginConfig = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
    var pageCallCount = 0;
    var measurementIds = getIds(pluginConfig.measurementIds);

    var initConfig = _objectSpread(_objectSpread({}, defaultConfig), pluginConfig);

    return {
      name: 'google-analytics',
      config: initConfig,
      // Load gtag. js and define gtag
      initialize: function initialize(_ref) {
        var config = _ref.config,
            instance = _ref.instance;
        var dataLayerName = config.dataLayerName,
            customScriptSrc = config.customScriptSrc,
            gtagName = config.gtagName,
            gtagConfig = config.gtagConfig,
            debug = config.debug,
            nonce = config.nonce;
        /* Inject google gtag. js script if not found */

        /* If other gtags are loaded already, add ours anyway */

        var customLayerName = dataLayerName ?  "&l=". concat(dataLayerName) : "";
        var src = customScriptSrc || "". concat(gtagScriptSource, "?id=").concat(measurementIds[0]).concat(customLayerName);

        if (! scriptLoaded(src)) {
          var script = document.createElement('script');
          script.async = true;
          script.src = src;

          if (nonce) {
            script.setAttribute('nonce', nonce);
          }

          document.body.appendChild(script);
        }
        /* Set up gtag and datalayer */


        if (! window[dataLayerName]) {
          window[dataLayerName] = window[dataLayerName] || [];
        }

        if (!window[gtagName]) {
          window[gtagName] = function () {
            window[dataLayerName].push(arguments);
          };
        }

        window[gtagName]('js', new Date()); // Initialize tracker instances on page

        var gtagConf = _objectSpread(_objectSpread({}, defaultGtagConf), gtagConfig ?  gtagConfig : {}); // You must explicitly delete the debug_mode parameter or all sessions will fire in debug more.  Setting it false is not enough. 
        // https://support.google.com/analytics/answer/7201382?hl=en&ref_topic=9303319#zippy=%2Cgoogle-tag-websites: ~: text=To%20disable%20debug%20mode%2C%20exclude%20the%20%27debug_mode%27%20parameter%3B%20setting%20the%20parameter%20to%20false%20doesn%27t%20disable%20debug%20mode. 


        if (debug === true) {
          gtagConf.debug_mode = true;
        } else {
          delete gtagConf.debug_mode;
        }
        /* set custom dimensions from user traits */


        var user = instance.user() || {};
        var traits = user.traits || {};

        if (Object.keys(traits).length) {
          window[gtagName]('set', 'user_properties', traits);
        }
        /* Initialize all measurementIds */


        for (var i = 0; i < measurementIds.length; i++) {
          if (! loadedInstances[measurementIds[i]]) {
            window[gtagName]('config', measurementIds[i], gtagConf);
            loadedInstances[measurementIds[i]] = true;
          }
        }
      },
      // Set parameter scope at user level with 'set' method
      identify: function identify(_ref2) {
        var payload = _ref2.payload,
            config = _ref2.config;
        var gtagName = config.gtagName;
        if (! window[gtagName] || !measurementIds.length) return;

        // Gebruik gehashte user ID gebaseerd op email + SharePoint ID
        var hashedUserId = getHashedUserId();
        
        if (hashedUserId) {
          // https://developers.google.com/analytics/devguides/collection/ga4/user-id? platform=websites#send_user_ids
          window[gtagName]('set', {
            user_id: hashedUserId
          });
          console.log('Set hashed user_id:', hashedUserId);
        } else if (payload.userId) {
          // Fallback naar originele userId als hashing faalt
          window[gtagName]('set', {
            user_id: payload.userId
          });
          console.log('Set userid (fallback):', payload.userId);
        }
        
        // TODO verify this
        // https://developers.google.com/analytics/devguides/collection/ga4/user-properties?technology=websites


        if (Object.keys(payload. traits).length) {
          /* gtag('set', 'user_properties', {
            favorite_composer: 'Mahler',
            favorite_instrument: 'double bass',
            season_ticketholder: 'true'
          }) */
          window[gtagName]('set', 'user_properties', payload.traits); // console.log('Set userprops', payload.traits)
        }
      },
      // Set parameter scope at page level with 'config' method
      page: function page(_ref3) {
        var payload = _ref3.payload,
            config = _ref3.config,
            instance = _ref3.instance;
        var gtagName = config. gtagName,
            gtagConfig = config. gtagConfig;
        if (!window[gtagName] || !measurementIds. length) return;
        var properties = payload.properties;
        var send_to = properties. send_to;
        var campaign = instance.getState('context. campaign'); // console.log('ga page properties', properties)

        /* Create pageview-related properties */

        var pageView = {
          page_title:  properties.title,
          page_location: properties.url,
          page_path: properties.path || document.location.pathname,
          page_hash: properties.hash,
          page_search: properties.page_search,
          page_referrer: properties.referrer
        };
        var campaignData = addCampaignData(campaign);
        
        // Gebruik gehashte user ID in plaats van originele userId
        var hashedUserId = getHashedUserId();
        var userId = hashedUserId || instance.user('userId');

        var finalPayload = _objectSpread(_objectSpread(_objectSpread(_objectSpread({}, send_to ?  {
          send_to: send_to
        } : {}), pageView), campaignData), userId ?  {
          user_id: userId
        } : {});
        /* If send_page_view true, ignore first analytics.page call */


        if (gtagConfig && gtagConfig.send_page_view && pageCallCount === 0) {
          pageCallCount++; // console.log('ignore first pageCallCount', pageCallCount)

          return;
        } // console.log('Send page_view payload', finalPayload)


        window[gtagName]('event', 'page_view', finalPayload); // Set after initial page view

        pageCallCount++;
      },
      // Set parameter scope at event level with 'event' method
      track: function track(_ref4) {
        var payload = _ref4.payload,
            config = _ref4.config,
            instance = _ref4.instance;
        var properties = payload.properties,
            event = payload.event;
        var campaign = instance.getState('context. campaign');
        var gtagName = config.gtagName;
        if (!window[gtagName] || !measurementIds.length) return;
        var campaignData = addCampaignData(campaign);
        
        // Gebruik gehashte user ID in plaats van originele userId
        var hashedUserId = getHashedUserId();
        var userId = hashedUserId || instance.user('userId');
        
        // Limits https://support.google.com/analytics/answer/9267744

        var finalPayload = _objectSpread(_objectSpread(_objectSpread({}, properties), campaignData), userId ? {
          user_id:  userId
        } : {});
        /*
          console.log('finalPayload', finalPayload)
          console.log('event', event)
        */

        /* Send data to Google Analytics
          Signature gtag('event', '<event_name>', {
            <event_params>key: value,
          })
        */


        window[gtagName]('event', event, finalPayload);
      },

      /* Verify gtag loaded and ready to use */
      loaded: function loaded() {
        var dataLayerName = initConfig.dataLayerName,
            customScriptSrc = initConfig.customScriptSrc;
        var hasDataLayer = dataLayerName && window[dataLayerName] && Array.prototype.push === window[dataLayerName].push;
        return scriptLoaded(customScriptSrc || gtagScriptSource) && hasDataLayer;
      },

      /* Custom methods */
      methods: {
        addTag: function addTag(tagId) {
          var settings = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};

          // https://developers.google.com/tag-platform/devguides/install-gtagjs#add_products_to_your_tag
          if (window[initConfig.gtagName]) {
            window[initConfig.gtagName]('config', tagId, settings); // Add tag id

            if (measurementIds && ! measurementIds.includes(tagId)) {
              measurementIds = measurementIds.concat(tagId);
            }
          }
        },

        /* Disable gtag for user */
        disable: function disable(ids) {
          var gaIds = ids ? getIds(ids) : measurementIds;

          for (var i = 0; i < measurementIds.length; i++) {
            var gaId = measurementIds[i];

            if (gaIds.includes(gaId)) {
              // https://developers.google.com/analytics/devguides/collection/gtagjs/user-opt-out
              window["ga-disable-". concat(gaId)] = true;
            }
          }
        },

        /* Enable gtag for user */
        enable: function enable(ids) {
          var gaIds = ids ? getIds(ids) : measurementIds;

          for (var i = 0; i < measurementIds.length; i++) {
            var gaId = measurementIds[i];

            if (gaIds.includes(gaId)) {
              // https://developers.google.com/analytics/devguides/collection/gtagjs/user-opt-out
              window["ga-disable-".concat(gaId)] = false;
            }
          }
        },

        /* Get current SharePoint user info */
        getUserInfo: function getUserInfo() {
          return getCurrentUserInfo();
        },

        /* Get hashed user ID */
        getHashedUserId: function getHashedUserIdMethod() {
          return getHashedUserId();
        }
      }
    };
  }

  function getIds(measurementIds) {
    if (!measurementIds) throw new Error('No GA Measurement ID defined');

    if (Array.isArray(measurementIds)) {
      return measurementIds;
    }

    if (typeof measurementIds === 'string') {
      return [measurementIds];
    }

    throw new Error('GA Measurement ID must be string or array of strings');
  }
  /**
   * Add campaign data to GA payload https://bit.ly/34qFCPn
   * @param {Object} [campaignData={}] [description]
   * @param {String} [campaignData.campaignName] - Name of campaign
   * @param {String} [campaignData. campaignSource] - Source of campaign
   * @param {String} [campaignData.campaignMedium] - Medium of campaign
   * @param {String} [campaignData.campaignContent] - Content of campaign
   * @param {String} [campaignData.campaignKeyword] - Keyword of campaign
   */


  function addCampaignData() {
    var campaignData = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
    var campaign = {};
    var id = campaignData.id,
        name = campaignData.name,
        source = campaignData. source,
        medium = campaignData.medium,
        content = campaignData.content,
        keyword = campaignData.keyword;
    if (id) campaign.campaignId = id;
    if (name) campaign.campaignName = name;
    if (source) campaign.campaignSource = source;
    if (medium) campaign.campaignMedium = medium;
    if (content) campaign.campaignContent = content;
    if (keyword) campaign.campaignKeyword = keyword;
    return campaign;
  }

  function scriptLoaded(scriptSrc) {
    var scripts = document.querySelectorAll('script[src]');
    var regex = new RegExp("^". concat(scriptSrc));
    return Boolean(Object.values(scripts).filter(function (value) {
      return regex.test(value.src);
    }).length);
  }

  /* This module will shake out unused code + work in browser and node 🎉 */

  var index = googleAnalytics ;
  /* init for CDN usage. globalName. init() */

  var init = googleAnalytics ;

  exports["default"] = index;
  exports.init = init;

  Object.defineProperty(exports, '__esModule', { value: true });

  return exports;

})({});