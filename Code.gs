/**
 * ============================================================================
 * DESIMURGA™ — FARM-TO-HOME ECOMMERCE ERP
 * Farm → Fresh → Your Home
 * ----------------------------------------------------------------------------
 * Code.gs — Complete server-side application.
 * Google Apps Script V8 runtime. Google Sheets as relational database.
 * ============================================================================
 */

/* ============================================================================
 * SECTION 0 — CONSTANTS, SCHEMA, ENUMS
 * ========================================================================== */

var APP_NAME = 'DesiMurga';

var PROP_KEYS = {
  SPREADSHEET_ID: 'DM_SPREADSHEET_ID',
  ROOT_FOLDER_ID: 'DM_ROOT_FOLDER_ID',
  INVOICE_FOLDER_ID: 'DM_INVOICE_FOLDER_ID',
  PRODUCT_IMAGE_FOLDER_ID: 'DM_PRODUCT_IMAGE_FOLDER_ID',
  REVIEW_PHOTO_FOLDER_ID: 'DM_REVIEW_PHOTO_FOLDER_ID',
  DOCUMENTS_FOLDER_ID: 'DM_DOCUMENTS_FOLDER_ID',
  TRIGGERS_INITIALIZED: 'DM_TRIGGERS_INITIALIZED'
};

// Full relational schema. 35 logical tables per specification.
var SCHEMA = {
  CONFIG: ['Key', 'Value', 'UpdatedAt'],
  USERS: ['UserId', 'Email', 'Name', 'Role', 'Active', 'CreatedAt'],
  ROLES: ['RoleId', 'RoleName', 'Description'],
  AUDIT_LOG: ['LogId', 'Timestamp', 'User', 'Role', 'Action', 'Entity', 'EntityId', 'OldValue', 'NewValue', 'Description'],
  SEQUENCES: ['SequenceKey', 'LastValue'],

  CUSTOMERS: ['CustomerId', 'Name', 'Mobile', 'Email', 'CreatedAt', 'Active', 'PhotoUrl'],
  CUSTOMER_ADDRESSES: ['AddressId', 'CustomerId', 'Label', 'AddressLine', 'Area', 'PIN', 'Lat', 'Lng', 'IsDefault', 'CreatedAt'],
  REVIEWS: ['ReviewId', 'CustomerId', 'OrderId', 'ProductId', 'Rating', 'Comment', 'PhotoUrl', 'CreatedAt', 'VerifiedPurchase', 'AdminApproved', 'Published'],

  PRODUCTS: ['ProductId', 'Name', 'Description', 'Category', 'ImageUrl', 'Images', 'Price', 'Unit', 'EstWeight', 'MinQty', 'MaxQty', 'Available', 'Featured', 'Active', 'BreedType', 'BirdType', 'Source', 'AgeRange', 'TypicalWeight', 'FarmingMethod', 'CreatedAt'],
  PRODUCT_VARIANTS: ['VariantId', 'ProductId', 'VariantName', 'WeightSlabMin', 'WeightSlabMax', 'PriceOverride', 'Active'],
  PRICE_RULES: ['RuleId', 'ProductId', 'RuleType', 'Value', 'ValidFrom', 'ValidTo', 'Active'],

  SUPPLIERS: ['SupplierId', 'BusinessName', 'ContactPerson', 'Phone', 'WhatsApp', 'Address', 'Location', 'BreedType', 'MinOrder', 'Price', 'AvgWeight', 'Capacity', 'LeadTime', 'Reliability', 'Active', 'Notes', 'CreatedAt'],
  SUPPLIER_PRODUCTS: ['Id', 'SupplierId', 'ProductType', 'Price', 'MinOrder', 'Active'],
  PURCHASE_ORDERS: ['PoId', 'SupplierId', 'Date', 'RequiredDate', 'Product', 'Quantity', 'TargetWeight', 'AgreedPrice', 'EstimatedTotal', 'Status', 'Notes', 'CreatedAt'],
  PURCHASE_ORDER_LINES: ['LineId', 'PoId', 'ProductType', 'Quantity', 'Price'],
  RECEIVING: ['ReceivingId', 'PoId', 'QtyOrdered', 'QtyReceived', 'Rejected', 'Mortality', 'Accepted', 'AvgLiveWeight', 'TotalLiveWeight', 'ReceivingDate', 'QcStatus'],

  LIVE_BATCHES: ['BatchId', 'SupplierId', 'ReceivingId', 'BirdType', 'Sex', 'Breed', 'QtyReceived', 'QtyAvailable', 'QtyReserved', 'QtyProcessing', 'QtyProcessed', 'Mortality', 'AvgLiveWeight', 'TotalLiveWeight', 'ArrivalDate', 'ExpectedProcessingDate', 'Status'],
  LIVE_INVENTORY: ['Id', 'TotalLiveBirds', 'Available', 'Reserved', 'Processing', 'Processed', 'Mortality', 'TotalLiveWeight', 'LastUpdated'],
  QC_RECORDS: ['QcId', 'BatchId', 'SupplierId', 'ReceivedQty', 'Healthy', 'Rejected', 'Mortality', 'Condition', 'WeightCheck', 'Notes', 'Inspector', 'Timestamp', 'Status'],
  MORTALITY_LOG: ['Id', 'BatchId', 'Qty', 'Date', 'Reason', 'RecordedBy'],

  PROCESSING_BATCHES: ['ProcBatchId', 'SourceLiveBatchId', 'BirdCount', 'TotalLiveWeight', 'BirdsProcessed', 'Rejected', 'Mortality', 'DressedWeight', 'SaleableWeight', 'Waste', 'YieldPct', 'SaleableYieldPct', 'Status', 'Date'],
  DRESSED_INVENTORY: ['Id', 'ProcBatchId', 'ProductId', 'Cut', 'Weight', 'AvailableWeight', 'ReservedWeight', 'SoldWeight', 'Wastage', 'StorageStatus', 'CreatedAt'],
  PACKAGING: ['PackId', 'OrderId', 'ProductId', 'Weight', 'PackedBy', 'PackedAt'],

  ORDERS: ['OrderId', 'CustomerId', 'Status', 'PaymentStatus', 'Subtotal', 'DeliveryFee', 'Discount', 'Total', 'AddressId', 'DeliverySlotId', 'SpecialInstruction', 'CreatedAt', 'UpdatedAt'],
  ORDER_LINES: ['LineId', 'OrderId', 'ProductId', 'Qty', 'EstWeight', 'Price', 'LineTotal'],
  PAYMENTS: ['PaymentId', 'OrderId', 'Amount', 'Method', 'UpiRef', 'Timestamp', 'Status', 'VerificationStatus'],
  INVOICES: ['InvoiceId', 'OrderId', 'PdfUrl', 'CreatedAt'],

  DELIVERY_ZONES: ['ZoneId', 'Name', 'CenterLat', 'CenterLng', 'RadiusKm', 'Active'],
  DELIVERIES: ['DeliveryId', 'OrderId', 'Status', 'SlotId', 'AssignedTo', 'DeliveredAt', 'Notes'],
  DELIVERY_SLOTS: ['SlotId', 'Name', 'StartTime', 'EndTime', 'Capacity', 'Booked', 'Active'],

  NOTIFICATIONS: ['NotifId', 'Event', 'Recipient', 'Channel', 'Payload', 'Status', 'CreatedAt'],

  COUPONS: ['CouponId', 'Code', 'DiscountType', 'Value', 'MinOrder', 'ValidFrom', 'ValidTo', 'Active', 'UsageLimit', 'UsedCount'],
  LOYALTY: ['Id', 'CustomerId', 'Points', 'UpdatedAt'],

  EXPENSES: ['ExpenseId', 'Category', 'Amount', 'Date', 'Notes'],
  DAILY_SUMMARY: ['Date', 'Sales', 'Orders', 'Paid', 'Pending', 'Cancelled', 'AvgOrderValue'],

  WALLET_TRANSACTIONS: ['TxnId', 'CustomerId', 'Type', 'Amount', 'Reason', 'RefId', 'Balance', 'CreatedAt'],
  REFERRALS: ['ReferralId', 'ReferrerCustomerId', 'ReferralCode', 'RefereeCustomerId', 'Status', 'RewardAmount', 'CreatedAt'],
  GIFT_CARDS: ['CardId', 'Code', 'InitialValue', 'Balance', 'IssuedTo', 'ExpiresAt', 'Active', 'CreatedAt'],
  SUBSCRIPTIONS: ['SubscriptionId', 'CustomerId', 'ProductId', 'Qty', 'Frequency', 'NextOrderDate', 'AddressId', 'SlotId', 'Status', 'CreatedAt', 'LastOrderId'],
  WHOLESALE_ACCOUNTS: ['AccountId', 'CustomerId', 'BusinessName', 'GSTIN', 'CreditLimit', 'CreditUsed', 'PaymentTerms', 'DiscountPercent', 'Active', 'CreatedAt'],
  RETURNS: ['ReturnId', 'OrderId', 'CustomerId', 'Reason', 'RefundMethod', 'RefundAmount', 'Status', 'RequestedAt', 'ResolvedAt'],
  SUPPORT_TICKETS: ['TicketId', 'CustomerId', 'OrderId', 'Subject', 'Category', 'Priority', 'Status', 'CreatedAt', 'UpdatedAt'],
  TICKET_MESSAGES: ['MessageId', 'TicketId', 'Sender', 'SenderRole', 'Message', 'Timestamp'],
  CAMPAIGNS: ['CampaignId', 'Name', 'Segment', 'Channel', 'MessageTemplate', 'Status', 'ScheduledAt', 'SentCount', 'CreatedAt'],
  VENDOR_PAYMENTS: ['PaymentId', 'SupplierId', 'PoId', 'Amount', 'Method', 'Status', 'PaidAt', 'Notes'],
  WISHLIST: ['WishlistId', 'CustomerId', 'ProductId', 'CreatedAt'],
  STATIC_PAGES: ['PageKey', 'Title', 'Content', 'UpdatedAt']
};

var ORDER_STATUSES = ['DRAFT', 'PENDING_PAYMENT', 'PAID', 'CONFIRMED', 'RESERVED', 'PROCESSING', 'PACKED', 'READY_FOR_DELIVERY', 'OUT_FOR_DELIVERY', 'DELIVERED', 'COMPLETED', 'CANCELLED', 'REFUNDED'];
var PAYMENT_STATUSES = ['PENDING', 'PAID', 'FAILED', 'REFUNDED', 'PARTIAL', 'MANUAL_VERIFIED'];
var INVENTORY_STATUSES = ['ACTIVE', 'DEPLETED', 'CLOSED'];
var DELIVERY_STATUSES = ['PENDING', 'ASSIGNED', 'OUT_FOR_DELIVERY', 'DELIVERED', 'FAILED'];
var PRODUCT_CATEGORIES = ['Whole Desi Hen', 'Curry Cut', 'Premium Curry Cut', 'Family Pack'];
var ROLE_LIST = ['ADMIN', 'MANAGER', 'PROCUREMENT', 'FARM', 'PROCESSING', 'PACKING', 'DELIVERY', 'CUSTOMER'];

var ORDER_TRANSITIONS = {
  DRAFT: ['PENDING_PAYMENT', 'PAID', 'CONFIRMED', 'RESERVED', 'PROCESSING', 'PACKED', 'CANCELLED'],
  PENDING_PAYMENT: ['PAID', 'CONFIRMED', 'RESERVED', 'PROCESSING', 'PACKED', 'CANCELLED'],
  PAID: ['CONFIRMED', 'RESERVED', 'PROCESSING', 'PACKED', 'REFUNDED', 'CANCELLED'],
  CONFIRMED: ['RESERVED', 'PROCESSING', 'PACKED', 'CANCELLED'],
  RESERVED: ['PROCESSING', 'PACKED', 'CANCELLED'],
  PROCESSING: ['PACKED', 'READY_FOR_DELIVERY', 'CANCELLED'],
  PACKED: ['READY_FOR_DELIVERY', 'OUT_FOR_DELIVERY', 'CANCELLED'],
  READY_FOR_DELIVERY: ['OUT_FOR_DELIVERY', 'DELIVERED', 'CANCELLED'],
  OUT_FOR_DELIVERY: ['DELIVERED', 'FAILED'],
  DELIVERED: ['COMPLETED'],
  COMPLETED: [],
  CANCELLED: [],
  REFUNDED: []
};

var DEFAULT_CONFIG = {
  brandName: 'DesiMurga™',
  tagline: 'Farm → Fresh → Your Home',
  locationName: 'Boalia, Garia, Kolkata',
  farmAddress: 'Boalia, Garia, Kolkata, West Bengal, India',
  businessAddress: 'Boalia, Garia, Kolkata, West Bengal, India',
  deliveryCenterLat: '22.4402',
  deliveryCenterLng: '88.3897',
  deliveryRadiusKm: '5',
  minimumLiveInventory: '50',
  maximumLiveInventory: '100',
  safetyStock: '10',
  currency: 'INR',
  currencySymbol: '₹',
  timezone: 'Asia/Kolkata',
  supportPhone: '+91-90000-00000',
  supportEmail: 'support@desimurga.example',
  whatsappNumber: '+91-90000-00000',
  upiId: 'desimurga@upi',
  upiQrUrl: '',
  youtubeFarmLiveId: '',
  youtubeProcessingLiveId: '',
  minimumOrder: '199',
  deliveryFee: '30',
  freeDeliveryThreshold: '499',
  taxPercent: '0',
  invoicePrefix: 'INV',
  orderPrefix: 'DM',
  logoUrl: '',
  faviconUrl: '',
  loyaltyEarnRatePercent: '2',
  loyaltyPointValue: '1',
  loyaltyMinRedeem: '50',
  referralRewardAmount: '50',
  refereeDiscountAmount: '50',
  wholesaleDefaultDiscountPercent: '10',
  supportTicketSlaHours: '24'
};

var SUBSCRIPTION_FREQUENCIES = ['WEEKLY', 'BIWEEKLY', 'MONTHLY'];
var SUBSCRIPTION_STATUSES = ['ACTIVE', 'PAUSED', 'CANCELLED'];
var RETURN_STATUSES = ['REQUESTED', 'APPROVED', 'REJECTED', 'REFUNDED'];
var TICKET_STATUSES = ['OPEN', 'IN_PROGRESS', 'RESOLVED', 'CLOSED'];
var TICKET_CATEGORIES = ['Order Issue', 'Quality Concern', 'Delivery Delay', 'Payment Issue', 'General Inquiry'];
var CAMPAIGN_SEGMENTS = ['ALL_CUSTOMERS', 'REPEAT_CUSTOMERS', 'CHURN_RISK', 'HIGH_VALUE', 'WHOLESALE'];

/* ============================================================================
 * SECTION 1 — LOW-LEVEL SPREADSHEET / PROPERTY HELPERS
 * ========================================================================== */

function getSpreadsheetId_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(PROP_KEYS.SPREADSHEET_ID);
}

function ss_() {
  var id = getSpreadsheetId_();
  if (!id) {
    throw new Error('Database not initialized. Run initializeDatabase() first.');
  }
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) {
  var sheet = ss_().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet not found: ' + name + '. Run repairDatabase().');
  }
  return sheet;
}

function sheetHeaders_(name) {
  return SCHEMA[name];
}

function sheetToObjects_(name) {
  var sheet = getSheet_(name);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  
  var headers = sheetHeaders_(name);
  var values = sheet.getRange(2, 1, lastRow - 1, Math.min(lastCol, headers.length)).getValues();
  var out = [];
  
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row || row.join('') === '') continue;
    var obj = { _row: i + 2 };
    
    for (var c = 0; c < headers.length; c++) {
      var val = row[c];
      
      // Convert native Date objects to ISO string to prevent null responses
      if (val instanceof Date) {
        val = val.toISOString();
      }
      
      obj[headers[c]] = (val === undefined || val === null) ? '' : val;
    }
    out.push(obj);
  }
  return out;
}

function objectToRowArray_(name, obj) {
  var headers = sheetHeaders_(name);
  return headers.map(function (h) { return (obj[h] === undefined || obj[h] === null) ? '' : obj[h]; });
}

function appendRow_(name, obj) {
  var sheet = getSheet_(name);
  sheet.appendRow(objectToRowArray_(name, obj));
  SpreadsheetApp.flush();
  return obj;
}

function findRowIndexById_(name, idField, idValue) {
  var sheet = getSheet_(name);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var headers = sheetHeaders_(name);
  var idCol = headers.indexOf(idField) + 1;
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idValue)) return i + 2;
  }
  return -1;
}

function updateRowFields_(name, idField, idValue, fields) {
  var sheet = getSheet_(name);
  var rowIndex = findRowIndexById_(name, idField, idValue);
  if (rowIndex === -1) throw new Error(name + ' record not found: ' + idValue);
  var headers = sheetHeaders_(name);
  Object.keys(fields).forEach(function (key) {
    var col = headers.indexOf(key);
    if (col === -1) return;
    sheet.getRange(rowIndex, col + 1).setValue(fields[key]);
  });
  SpreadsheetApp.flush();
  return true;
}

function deleteRow_(name, idField, idValue) {
  var sheet = getSheet_(name);
  var rowIndex = findRowIndexById_(name, idField, idValue);
  if (rowIndex === -1) return false;
  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  return true;
}

function getRowById_(name, idField, idValue) {
  var rowIndex = findRowIndexById_(name, idField, idValue);
  if (rowIndex === -1) return null;
  var sheet = getSheet_(name);
  var headers = sheetHeaders_(name);
  var values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  var obj = { _row: rowIndex };
  for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[c];
  return obj;
}

/* ============================================================================
 * SECTION 2 — DEFENSIVE PARSING & FORMATTING UTILITIES
 * ========================================================================== */

function safeString_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function safeNumber_(v, fallback) {
  var n = Number(v);
  if (isNaN(n)) return (fallback === undefined ? 0 : fallback);
  return n;
}

function safeBoolean_(v, fallback) {
  if (v === null || v === undefined) return fallback === true;
  if (typeof v === 'boolean') return v;
  var str = String(v).trim().toUpperCase();
  if (str === 'TRUE' || str === '1' || str === 'YES') return true;
  if (str === 'FALSE' || str === '0' || str === 'NO') return false;
  return fallback === true;
}

function safeDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var d = new Date(v);
  if (isNaN(d.getTime())) return new Date();
  return d;
}

function esc_(str) {
  return safeString_(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function nowIso_() {
  return new Date().toISOString();
}

function todayStamp_() {
  return Utilities.formatDate(new Date(), getConfigValue_('timezone', 'Asia/Kolkata'), 'yyyyMMdd');
}

function formatCurrency(n) { 
  var sym = '₹';
  try {
    var map = getConfigMap_();
    if (map && map.currencySymbol) {
      sym = map.currencySymbol;
    }
  } catch (e) {}
  
  var val = parseFloat(n);
  if (isNaN(val)) val = 0;
  return sym + val.toFixed(2); 
}

/**
 * Universal Google Drive File ID Extractor
 * Accepts drive view URLs, open URLs, direct IDs, lh3 links, etc.
 */
function extractDriveId_(url) {
  if (!url) return '';
  var str = String(url).trim();
  var fileId = '';
  
  if (str.indexOf('/file/d/') !== -1) {
    fileId = str.split('/file/d/')[1].split('/')[0].split('?')[0];
  } else if (str.indexOf('lh3.googleusercontent.com/d/') !== -1) {
    fileId = str.split('lh3.googleusercontent.com/d/')[1].split('?')[0].split('=')[0];
  } else if (str.indexOf('id=') !== -1) {
    fileId = str.split('id=')[1].split('&')[0];
  } else if (str.length > 20 && str.indexOf('/') === -1 && str.indexOf('.') === -1) {
    fileId = str;
  }
  return fileId;
}

/* ============================================================================
 * SECTION 3 — RESPONSE STANDARD & ERROR HANDLING
 * ========================================================================== */

function successResponse_(data, message) {
  return { success: true, data: data === undefined ? {} : data, message: message || '', timestamp: nowIso_() };
}

function errorResponse_(message, code) {
  return { success: false, code: code || 'ERROR', message: message || 'An error occurred.', timestamp: nowIso_() };
}

function handleServerError_(error, context) {
  try {
    Logger.log('ERROR in ' + context + ': ' + (error && error.stack ? error.stack : error));
    writeAudit_('SYSTEM', 'SYSTEM', 'SERVER_ERROR', context, '', '', String(error), context + ' failed');
  } catch (ignored) {}
  return errorResponse_('Something went wrong while processing your request. Please try again.', 'SERVER_ERROR');
}

function safeInvoke_(context, fn) {
  try {
    return fn();
  } catch (err) {
    return handleServerError_(err, context);
  }
}

/* ============================================================================
 * SECTION 4 — SEQUENCE / ID GENERATOR
 * ========================================================================== */

function nextSequenceValue_(key) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_('SEQUENCES');
    var lastRow = sheet.getLastRow();
    var rowIndex = -1;
    var next = 1;
    if (lastRow >= 2) {
      var keys = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) { rowIndex = i + 2; next = Number(keys[i][1]) + 1; break; }
      }
    }
    if (rowIndex === -1) {
      sheet.appendRow([key, next]);
    } else {
      sheet.getRange(rowIndex, 2).setValue(next);
    }
    return next;
  } finally {
    lock.releaseLock();
  }
}

function pad_(n, width) {
  var s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

function generateId_(type) {
  switch (type) {
    case 'CUS': return 'CUS-' + pad_(nextSequenceValue_('CUS'), 6);
    case 'PRD': return 'PRD-' + pad_(nextSequenceValue_('PRD'), 6);
    case 'SUP': return 'SUP-' + pad_(nextSequenceValue_('SUP'), 6);
    case 'ADDR': return 'ADDR-' + pad_(nextSequenceValue_('ADDR'), 6);
    case 'REVIEW': return 'REV-' + pad_(nextSequenceValue_('REVIEW'), 6);
    case 'BATCH': return 'BATCH-' + todayStamp_() + '-' + pad_(nextSequenceValue_('BATCH-' + todayStamp_()), 3);
    case 'PO': return 'PO-' + todayStamp_() + '-' + pad_(nextSequenceValue_('PO-' + todayStamp_()), 3);
    case 'RECV': return 'RECV-' + todayStamp_() + '-' + pad_(nextSequenceValue_('RECV-' + todayStamp_()), 3);
    case 'QC': return 'QC-' + todayStamp_() + '-' + pad_(nextSequenceValue_('QC-' + todayStamp_()), 3);
    case 'PROC': return 'PROC-' + todayStamp_() + '-' + pad_(nextSequenceValue_('PROC-' + todayStamp_()), 3);
    case 'DRESSED': return 'DRS-' + pad_(nextSequenceValue_('DRESSED'), 6);
    case 'DM': return (getConfigValue_('orderPrefix', 'DM')) + '-' + todayStamp_() + '-' + pad_(nextSequenceValue_('DM-' + todayStamp_()), 4);
    case 'INV': return (getConfigValue_('invoicePrefix', 'INV')) + '-' + todayStamp_() + '-' + pad_(nextSequenceValue_('INV-' + todayStamp_()), 4);
    case 'DEL': return 'DEL-' + todayStamp_() + '-' + pad_(nextSequenceValue_('DEL-' + todayStamp_()), 4);
    case 'PAY': return 'PAY-' + pad_(nextSequenceValue_('PAY'), 6);
    case 'NOTIF': return 'NTF-' + pad_(nextSequenceValue_('NOTIF'), 6);
    case 'COUPON': return 'CPN-' + pad_(nextSequenceValue_('COUPON'), 6);
    case 'EXP': return 'EXP-' + pad_(nextSequenceValue_('EXP'), 6);
    case 'PACK': return 'PACK-' + pad_(nextSequenceValue_('PACK'), 6);
    case 'ZONE': return 'ZONE-' + pad_(nextSequenceValue_('ZONE'), 3);
    case 'SLOT': return 'SLOT-' + pad_(nextSequenceValue_('SLOT'), 3);
    case 'USER': return 'USR-' + pad_(nextSequenceValue_('USER'), 5);
    case 'LOG': return 'LOG-' + pad_(nextSequenceValue_('LOG'), 8);
    case 'LOYALTY': return 'LOY-' + pad_(nextSequenceValue_('LOYALTY'), 6);
    case 'MORT': return 'MRT-' + pad_(nextSequenceValue_('MORT'), 6);
    case 'WALLET': return 'WTX-' + pad_(nextSequenceValue_('WALLET'), 6);
    case 'REF': return 'REF-' + pad_(nextSequenceValue_('REF'), 6);
    case 'GIFT': return 'GC-' + pad_(nextSequenceValue_('GIFT'), 6);
    case 'SUB': return 'SUB-' + pad_(nextSequenceValue_('SUB'), 6);
    case 'WHS': return 'WHS-' + pad_(nextSequenceValue_('WHS'), 5);
    case 'RET': return 'RET-' + pad_(nextSequenceValue_('RET'), 6);
    case 'TICKET': return 'TKT-' + pad_(nextSequenceValue_('TICKET'), 6);
    case 'TMSG': return 'TMS-' + pad_(nextSequenceValue_('TMSG'), 7);
    case 'CAMPAIGN': return 'CMP-' + pad_(nextSequenceValue_('CAMPAIGN'), 5);
    case 'VPAY': return 'VPY-' + pad_(nextSequenceValue_('VPAY'), 6);
    case 'WISH': return 'WSH-' + pad_(nextSequenceValue_('WISH'), 6);
    default: throw new Error('Unknown ID type: ' + type);
  }
}

/* ============================================================================
 * SECTION 5 — AUDIT LOG
 * ========================================================================== */

function writeAudit_(user, role, action, entity, entityId, oldValue, newValue, description) {
  try {
    appendRow_('AUDIT_LOG', {
      LogId: generateId_('LOG'),
      Timestamp: nowIso_(),
      User: user || 'SYSTEM',
      Role: role || 'SYSTEM',
      Action: action,
      Entity: entity,
      EntityId: entityId || '',
      OldValue: oldValue === undefined ? '' : String(oldValue),
      NewValue: newValue === undefined ? '' : String(newValue),
      Description: description || ''
    });
  } catch (e) {
    Logger.log('Audit log failure: ' + e);
  }
}

/* ============================================================================
 * SECTION 6 — CONFIGURATION SERVICE
 * ========================================================================== */

function getConfigMap_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('DM_CONFIG_MAP');
  if (cached) return JSON.parse(cached);
  var rows = sheetToObjects_('CONFIG');
  var map = {};
  rows.forEach(function (r) { map[r.Key] = r.Value; });
  cache.put('DM_CONFIG_MAP', JSON.stringify(map), 300);
  return map;
}

function getConfigValue_(key, fallback) {
  var map = getConfigMap_();
  if (map[key] === undefined || map[key] === '') return fallback;
  return map[key];
}

function setConfigValue_(key, value) {
  var rowIndex = findRowIndexById_('CONFIG', 'Key', key);
  if (rowIndex === -1) {
    appendRow_('CONFIG', { Key: key, Value: value, UpdatedAt: nowIso_() });
  } else {
    updateRowFields_('CONFIG', 'Key', key, { Value: value, UpdatedAt: nowIso_() });
  }
  CacheService.getScriptCache().remove('DM_CONFIG_MAP');
}

function getPublicConfig() {
  return safeInvoke_('getPublicConfig', function () {
    var map = getConfigMap_();
    var publicKeys = ['brandName', 'tagline', 'locationName', 'deliveryRadiusKm', 'currency', 'currencySymbol',
      'supportPhone', 'supportEmail', 'whatsappNumber', 'upiId', 'upiQrUrl', 'youtubeFarmLiveId',
      'youtubeProcessingLiveId', 'minimumOrder', 'deliveryFee', 'freeDeliveryThreshold', 'logoUrl', 'faviconUrl'];
    var out = {};
    publicKeys.forEach(function (k) { out[k] = map[k] !== undefined ? map[k] : DEFAULT_CONFIG[k]; });
    return successResponse_(out);
  });
}

function adminUpdateConfig(payload) {
  return safeInvoke_('adminUpdateConfig', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    Object.keys(payload || {}).forEach(function (key) {
      setConfigValue_(key, payload[key]);
    });
    
    CacheService.getScriptCache().remove('DM_CONFIG_MAP');
    writeAudit_(currentEmail_(), currentRole_(), 'CONFIG_UPDATED', 'CONFIG', '', '', JSON.stringify(payload), 'Settings updated');
    return successResponse_(getConfigMap_(), 'Settings saved successfully.');
  });
}

/* ============================================================================
 * SECTION 7 — AUTHENTICATION & AUTHORIZATION (SECURED)
 * ========================================================================== */

function currentEmail_() {
  try {
    // getActiveUser gets the REAL visitor's Google email (if authenticated)
    var email = Session.getActiveUser().getEmail();
    if (email) return email.toLowerCase().trim();
  } catch (e) {}
  return 'guest';
}

function currentRole_(adminToken) {
  var email = currentEmail_();
  
  // 1. Verify Google Account email against USERS table
  if (email !== 'guest') {
    var users = sheetToObjects_('USERS');
    for (var i = 0; i < users.length; i++) {
      if (String(users[i].Email || '').toLowerCase().trim() === email && safeBoolean_(users[i].Active)) {
        return users[i].Role;
      }
    }
  }

  // 2. Verify Session Admin Token (if user logged in via Admin PIN)
  if (adminToken) {
    var cache = CacheService.getScriptCache();
    var cachedRole = cache.get('DM_ADMIN_TOKEN_' + adminToken);
    if (cachedRole) return cachedRole;
  }

  return 'CUSTOMER';
}

function requireRole_(allowedRoles, adminToken) {
  var role = currentRole_(adminToken);
  if (allowedRoles.indexOf(role) === -1) {
    throw new Error('Access denied. Administrator authentication required.');
  }
  return role;
}

/**
 * Public bootstrap endpoint - strictly locks down isAdmin / isStaff for visitors
 */
function getBootstrapData(e) {
  return safeInvoke_('getBootstrapData', function () {
    var email = currentEmail_();
    var role = currentRole_();
    
    // STRICT SECURITY: Anonymous visitors are always CUSTOMER
    var isAdmin = (role === 'ADMIN');
    var isStaff = (ROLE_LIST.indexOf(role) !== -1 && role !== 'CUSTOMER');

    var cfg = getConfigMap_();
    var publicCfg = {};
    Object.keys(DEFAULT_CONFIG).forEach(function (k) { 
      publicCfg[k] = cfg[k] !== undefined && cfg[k] !== '' ? cfg[k] : DEFAULT_CONFIG[k]; 
    });
    
    var webAppUrl = '';
    try {
      webAppUrl = ScriptApp.getService().getUrl();
    } catch (err) {
      webAppUrl = '';
    }

    var initialProductId = '';
    if (e && e.parameter && e.parameter.p) {
      initialProductId = safeString_(e.parameter.p);
    }

    return successResponse_({
      email: email,
      role: role,
      isAdmin: isAdmin,
      isStaff: isStaff,
      config: publicCfg,
      webAppUrl: webAppUrl,
      initialProductId: initialProductId
    });
  });
}

/**
 * Secure Admin Login via Passcode / PIN (Stored in CONFIG sheet under 'adminPasscode')
 */
function verifyAdminLogin(passcode) {
  return safeInvoke_('verifyAdminLogin', function () {
    if (!passcode) return errorResponse_('Passcode is required.', 'AUTH_FAILED');
    
    var storedPasscode = getConfigValue_('adminPasscode', '123456'); // Default is 123456
    
    if (String(passcode).trim() === String(storedPasscode).trim()) {
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put('DM_ADMIN_TOKEN_' + token, 'ADMIN', 21600); // 6 hours session
      writeAudit_('ADMIN_LOGIN', 'ADMIN', 'LOGIN_SUCCESS', 'SYSTEM', '', '', '', 'Admin authenticated via Passcode');
      return successResponse_({ token: token, role: 'ADMIN' }, 'Admin authentication successful.');
    }
    
    writeAudit_('GUEST', 'GUEST', 'LOGIN_FAILED', 'SYSTEM', '', '', '', 'Failed Admin passcode attempt');
    return errorResponse_('Invalid Admin Passcode.', 'INVALID_CREDENTIALS');
  });
}

/* ============================================================================
 * SECTION 8 — DATABASE INITIALIZATION
 * ========================================================================== */

function initializeDatabase() {
  var report = { success: true, sheetsCreated: 0, sheetsExisting: 0, foldersCreated: 0, configurationInitialized: false };

  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(PROP_KEYS.SPREADSHEET_ID);
  var spreadsheet;
  if (ssId) {
    try { spreadsheet = SpreadsheetApp.openById(ssId); } catch (e) { ssId = null; }
  }
  if (!ssId) {
    spreadsheet = SpreadsheetApp.create('DesiMurga Database');
    ssId = spreadsheet.getId();
    props.setProperty(PROP_KEYS.SPREADSHEET_ID, ssId);
  }

  var sheetNames = Object.keys(SCHEMA);
  sheetNames.forEach(function (name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
      report.sheetsCreated++;
    } else {
      report.sheetsExisting++;
    }
    var headers = SCHEMA[name];
    var firstRowVals = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var hasHeaders = firstRowVals.join('') !== '';
    if (!hasHeaders) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#7a4a2b')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
    }
  });

  var defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    spreadsheet.deleteSheet(defaultSheet);
  }

  var rootFolderId = props.getProperty(PROP_KEYS.ROOT_FOLDER_ID);
  var rootFolder;
  if (rootFolderId) {
    try { rootFolder = DriveApp.getFolderById(rootFolderId); } catch (e) { rootFolderId = null; }
  }
  if (!rootFolderId) {
    rootFolder = DriveApp.createFolder('DesiMurga');
    props.setProperty(PROP_KEYS.ROOT_FOLDER_ID, rootFolder.getId());
    report.foldersCreated++;
  }
  var subFolders = [
    { key: PROP_KEYS.INVOICE_FOLDER_ID, name: 'Invoices' },
    { key: PROP_KEYS.PRODUCT_IMAGE_FOLDER_ID, name: 'ProductImages' },
    { key: PROP_KEYS.REVIEW_PHOTO_FOLDER_ID, name: 'ReviewPhotos' },
    { key: PROP_KEYS.DOCUMENTS_FOLDER_ID, name: 'Documents' }
  ];
  subFolders.forEach(function (sf) {
    var existingId = props.getProperty(sf.key);
    var exists = false;
    if (existingId) {
      try { DriveApp.getFolderById(existingId); exists = true; } catch (e) { exists = false; }
    }
    if (!exists) {
      var folder = rootFolder.createFolder(sf.name);
      props.setProperty(sf.key, folder.getId());
      report.foldersCreated++;
    }
  });

  var existingConfig = sheetToObjects_('CONFIG');
  var existingKeys = {};
  existingConfig.forEach(function (r) { existingKeys[r.Key] = true; });
  Object.keys(DEFAULT_CONFIG).forEach(function (key) {
    if (!existingKeys[key]) {
      appendRow_('CONFIG', { Key: key, Value: DEFAULT_CONFIG[key], UpdatedAt: nowIso_() });
    }
  });
  CacheService.getScriptCache().remove('DM_CONFIG_MAP');
  report.configurationInitialized = true;

  var existingRoles = sheetToObjects_('ROLES');
  if (existingRoles.length === 0) {
    var roleDescriptions = {
      ADMIN: 'Full system access', MANAGER: 'Operations & commerce management',
      PROCUREMENT: 'Supplier & purchase order management', FARM: 'Live inventory & farm operations',
      PROCESSING: 'Dressing / processing operations', PACKING: 'Order packing operations',
      DELIVERY: 'Delivery execution', CUSTOMER: 'Storefront customer'
    };
    ROLE_LIST.forEach(function (r, idx) {
      appendRow_('ROLES', { RoleId: 'ROLE-' + pad_(idx + 1, 2), RoleName: r, Description: roleDescriptions[r] });
    });
  }

  var zones = sheetToObjects_('DELIVERY_ZONES');
  if (zones.length === 0) {
    appendRow_('DELIVERY_ZONES', {
      ZoneId: generateId_('ZONE'), Name: 'Garia 5km Zone',
      CenterLat: DEFAULT_CONFIG.deliveryCenterLat, CenterLng: DEFAULT_CONFIG.deliveryCenterLng,
      RadiusKm: DEFAULT_CONFIG.deliveryRadiusKm, Active: true
    });
  }

  var slots = sheetToObjects_('DELIVERY_SLOTS');
  if (slots.length === 0) {
    var defaultSlots = [
      { Name: 'Morning', StartTime: '08:00', EndTime: '11:00' },
      { Name: 'Afternoon', StartTime: '12:00', EndTime: '15:00' },
      { Name: 'Evening', StartTime: '17:00', EndTime: '20:00' }
    ];
    defaultSlots.forEach(function (s) {
      appendRow_('DELIVERY_SLOTS', {
        SlotId: generateId_('SLOT'), Name: s.Name, StartTime: s.StartTime, EndTime: s.EndTime,
        Capacity: 25, Booked: 0, Active: true
      });
    });
  }

  var liveInv = sheetToObjects_('LIVE_INVENTORY');
  if (liveInv.length === 0) {
    appendRow_('LIVE_INVENTORY', {
      Id: 'SUMMARY', TotalLiveBirds: 0, Available: 0, Reserved: 0, Processing: 0,
      Processed: 0, Mortality: 0, TotalLiveWeight: 0, LastUpdated: nowIso_()
    });
  }

  var existingPages = sheetToObjects_('STATIC_PAGES');
  if (existingPages.length === 0) {
    var defaultPages = [
      { PageKey: 'about', Title: 'About Us',
        Content: 'DesiMurga sources mature Desi hens from verified local suppliers, keeps them in a small holding facility, and dresses each order fresh — no long cold storage. We believe fresh food should come with full transparency, which is why our Live Farm and Live Processing pages let you see exactly where your food comes from before you order.' },
      { PageKey: 'contact', Title: 'Contact Us',
        Content: 'We would love to hear from you. Reach us using the phone, WhatsApp, or email address shown in the footer of every page, or open a support ticket from any order in My Orders. Our team typically responds within one business day.' },
      { PageKey: 'faq', Title: 'Frequently Asked Questions',
        Content: 'Q: Do I need an account to order?\nA: No — just your name and mobile number at checkout.\n\nQ: Where do you deliver?\nA: Within our configured delivery radius of the farm location, shown on the checkout page.\n\nQ: How do I pay?\nA: Via UPI at checkout, or using wallet/gift card credit if you have any.\n\nQ: How fresh is the chicken?\nA: Birds are dressed to order the same day, not held in long-term cold storage.\n\nQ: Can I return an order?\nA: Yes — open "My Orders" and tap Return on any delivered order.' },
      { PageKey: 'terms', Title: 'Terms & Conditions',
        Content: 'By placing an order with DesiMurga you agree to pay the total shown at checkout, provide an accurate delivery address within our service area, and be available during your chosen delivery slot. Orders are prepared fresh to order and cannot be modified once processing has started.' },
      { PageKey: 'privacy', Title: 'Privacy Policy',
        Content: 'DesiMurga collects your name, mobile number, delivery address, and order history solely to fulfil and support your orders. We do not sell your information to third parties.' }
    ];
    defaultPages.forEach(function (p) {
      appendRow_('STATIC_PAGES', { PageKey: p.PageKey, Title: p.Title, Content: p.Content, UpdatedAt: nowIso_() });
    });
  }

  var users = sheetToObjects_('USERS');
  var hasAdmin = users.some(function (u) { return u.Role === 'ADMIN' && safeBoolean_(u.Active); });
  if (!hasAdmin) {
    var email = currentEmail_();
    var existingUserRow = users.filter(function (u) { return u.Email === email; })[0];
    if (existingUserRow) {
      updateRowFields_('USERS', 'UserId', existingUserRow.UserId, { Role: 'ADMIN', Active: true });
    } else {
      appendRow_('USERS', { UserId: generateId_('USER'), Email: email, Name: 'Administrator', Role: 'ADMIN', Active: true, CreatedAt: nowIso_() });
    }
  }

  writeAudit_(currentEmail_(), 'ADMIN', 'DATABASE_INITIALIZED', 'SYSTEM', '', '', '', 'initializeDatabase() executed');

  report.spreadsheetId = ssId;
  report.spreadsheetUrl = spreadsheet.getUrl();
  return report;
}

function adminInitializeDatabase() {
  return safeInvoke_('adminInitializeDatabase', function () {
    var report = initializeDatabase();
    return successResponse_(report, 'Database initialized successfully.');
  });
}

function initializeDemoData() {
  var report = { productsCreated: 0, suppliersCreated: 0 };
  var products = sheetToObjects_('PRODUCTS');
  if (products.length === 0) {
    var demoProducts = [
      { Name: 'Desi Hen — Whole', Category: 'Whole Desi Hen', Price: 420, Unit: 'per bird', EstWeight: 1.3, MinQty: 1, MaxQty: 5, ImageUrl: 'https://images.unsplash.com/photo-1548550023-2bdb3c5beed7?auto=format&fit=crop&w=600&q=80' },
      { Name: 'Desi Hen — Curry Cut', Category: 'Curry Cut', Price: 460, Unit: 'per kg', EstWeight: 1.0, MinQty: 1, MaxQty: 5, ImageUrl: 'https://images.unsplash.com/photo-1604503468506-a8da13d82791?auto=format&fit=crop&w=600&q=80' },
      { Name: 'Desi Hen — Small Family Pack', Category: 'Family Pack', Price: 850, Unit: 'per pack', EstWeight: 2.0, MinQty: 1, MaxQty: 3, ImageUrl: 'https://images.unsplash.com/photo-1587593810167-a84920ea0781?auto=format&fit=crop&w=600&q=80' },
      { Name: 'Desi Hen — Premium Cut', Category: 'Premium Curry Cut', Price: 520, Unit: 'per kg', EstWeight: 1.0, MinQty: 1, MaxQty: 5, ImageUrl: 'https://images.unsplash.com/photo-1604503468506-a8da13d82791?auto=format&fit=crop&w=600&q=80' }
    ];
    demoProducts.forEach(function (p) {
      appendRow_('PRODUCTS', {
        ProductId: generateId_('PRD'), Name: p.Name, Description: 'Farm-fresh ' + p.Name + ' from DesiMurga.',
        Category: p.Category, ImageUrl: p.ImageUrl, Price: p.Price, Unit: p.Unit, EstWeight: p.EstWeight,
        MinQty: p.MinQty, MaxQty: p.MaxQty, Available: true, Featured: true, Active: true,
        BreedType: 'Desi / Country', BirdType: 'Hen', Source: 'Verified local suppliers',
        AgeRange: '5-7 months', TypicalWeight: '1.1-1.5 kg live', FarmingMethod: 'Free-range',
        CreatedAt: nowIso_()
      });
      report.productsCreated++;
    });
  }
  var suppliers = sheetToObjects_('SUPPLIERS');
  if (suppliers.length === 0) {
    appendRow_('SUPPLIERS', {
      SupplierId: generateId_('SUP'), BusinessName: 'Garia Poultry Farm', ContactPerson: 'Ramesh Mondal',
      Phone: '+91-90000-11111', WhatsApp: '+91-90000-11111', Address: 'Garia, Kolkata', Location: 'Garia',
      BreedType: 'Desi', MinOrder: 20, Price: 250, AvgWeight: 1.3, Capacity: 100, LeadTime: 2,
      Reliability: 'Good', Active: true, Notes: 'Demo supplier record', CreatedAt: nowIso_()
    });
    report.suppliersCreated++;
  }
  return successResponse_(report, 'Demo data initialized.');
}

function adminInitializeDemoData() {
  return safeInvoke_('adminInitializeDemoData', function () {
    requireRole_(['ADMIN']);
    return initializeDemoData();
  });
}

/* ============================================================================
 * SECTION 9 — DATABASE HEALTH CHECK & REPAIR
 * ========================================================================== */

function runDatabaseHealthCheck() {
  var issues = [];
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(PROP_KEYS.SPREADSHEET_ID);
  if (!ssId) { issues.push('Spreadsheet not initialized.'); return { healthy: false, issues: issues }; }
  var spreadsheet;
  try { spreadsheet = SpreadsheetApp.openById(ssId); } catch (e) { issues.push('Spreadsheet ID invalid or inaccessible.'); return { healthy: false, issues: issues }; }

  Object.keys(SCHEMA).forEach(function (name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) { issues.push('Missing sheet: ' + name); return; }
    var headers = SCHEMA[name];
    var actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (actual[i] !== headers[i]) { issues.push('Header mismatch in ' + name + ' at column ' + (i + 1)); break; }
    }
  });

  [PROP_KEYS.ROOT_FOLDER_ID, PROP_KEYS.INVOICE_FOLDER_ID, PROP_KEYS.PRODUCT_IMAGE_FOLDER_ID, PROP_KEYS.REVIEW_PHOTO_FOLDER_ID, PROP_KEYS.DOCUMENTS_FOLDER_ID].forEach(function (key) {
    var id = props.getProperty(key);
    if (!id) { issues.push('Missing Drive folder property: ' + key); return; }
    try { DriveApp.getFolderById(id); } catch (e) { issues.push('Drive folder inaccessible: ' + key); }
  });

  var configRows = sheetToObjects_('CONFIG');
  if (configRows.length === 0) issues.push('CONFIG sheet is empty.');
  var roleRows = sheetToObjects_('ROLES');
  if (roleRows.length === 0) issues.push('ROLES sheet is empty.');

  ['PRODUCTS:ProductId', 'CUSTOMERS:CustomerId', 'ORDERS:OrderId', 'SUPPLIERS:SupplierId'].forEach(function (spec) {
    var parts = spec.split(':');
    var rows = sheetToObjects_(parts[0]);
    var seen = {};
    rows.forEach(function (r) {
      var id = r[parts[1]];
      if (!id) return;
      if (seen[id]) issues.push('Duplicate ID in ' + parts[0] + ': ' + id);
      seen[id] = true;
    });
  });

  return { healthy: issues.length === 0, issues: issues, checkedAt: nowIso_() };
}

function adminRunHealthCheck() {
  return safeInvoke_('adminRunHealthCheck', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(runDatabaseHealthCheck());
  });
}

function repairDatabase() {
  var report = initializeDatabase();
  return report;
}

function adminRepairDatabase() {
  return safeInvoke_('adminRepairDatabase', function () {
    requireRole_(['ADMIN']);
    var report = repairDatabase();
    return successResponse_(report, 'Database repaired.');
  });
}

/* ============================================================================
 * SECTION 10 — GEO / DELIVERY RADIUS VALIDATION
 * ========================================================================== */

function haversineKm_(lat1, lng1, lat2, lng2) {
  var R = 6371;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLng = (lng2 - lng1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function isWithinDeliveryRadius_(lat, lng) {
  var centerLat = safeNumber_(getConfigValue_('deliveryCenterLat', DEFAULT_CONFIG.deliveryCenterLat));
  var centerLng = safeNumber_(getConfigValue_('deliveryCenterLng', DEFAULT_CONFIG.deliveryCenterLng));
  var radius = safeNumber_(getConfigValue_('deliveryRadiusKm', DEFAULT_CONFIG.deliveryRadiusKm));
  if (!lat || !lng) return { withinRadius: false, distanceKm: null, reason: 'No coordinates supplied; manual verification required.' };
  var distance = haversineKm_(centerLat, centerLng, Number(lat), Number(lng));
  return { withinRadius: distance <= radius, distanceKm: Math.round(distance * 100) / 100, reason: distance <= radius ? '' : 'Address is outside the ' + radius + ' km service area.' };
}

function checkDeliveryRadius(lat, lng) {
  return safeInvoke_('checkDeliveryRadius', function () {
    return successResponse_(isWithinDeliveryRadius_(lat, lng));
  });
}

/* ============================================================================
 * SECTION 11 — PRODUCT CATALOG SERVICE
 * ========================================================================== */

function getShopProducts() {
  return safeInvoke_('getShopProducts', function () {
    var products = sheetToObjects_('PRODUCTS').filter(function (p) { 
      return safeBoolean_(p.Active, true) && safeBoolean_(p.Available, true); 
    });
    return successResponse_(products);
  });
}

function getAdminProducts() {
  return safeInvoke_('getAdminProducts', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(sheetToObjects_('PRODUCTS'));
  });
}

function adminSaveProduct(product) {
  return safeInvoke_('adminSaveProduct', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    if (!product || !safeString_(product.Name)) return errorResponse_('Product name is required.', 'VALIDATION_ERROR');
    var price = safeNumber_(product.Price);
    if (price <= 0) return errorResponse_('Product price must be greater than zero.', 'VALIDATION_ERROR');

    if (product.ProductId) {
      var existing = getRowById_('PRODUCTS', 'ProductId', product.ProductId);
      if (!existing) return errorResponse_('Product not found.', 'NOT_FOUND');
      updateRowFields_('PRODUCTS', 'ProductId', product.ProductId, {
        Name: safeString_(product.Name), Description: safeString_(product.Description),
        Category: safeString_(product.Category), ImageUrl: safeString_(product.ImageUrl),
        Images: safeString_(product.Images),
        Price: price, Unit: safeString_(product.Unit) || 'per unit',
        EstWeight: safeNumber_(product.EstWeight), MinQty: safeNumber_(product.MinQty, 1),
        MaxQty: safeNumber_(product.MaxQty, 10), Available: safeBoolean_(product.Available, true), Featured: safeBoolean_(product.Featured, false),
        Active: safeBoolean_(product.Active, true), BreedType: safeString_(product.BreedType), BirdType: safeString_(product.BirdType),
        Source: safeString_(product.Source), AgeRange: safeString_(product.AgeRange),
        TypicalWeight: safeString_(product.TypicalWeight), FarmingMethod: safeString_(product.FarmingMethod)
      });
      writeAudit_(currentEmail_(), currentRole_(), 'PRODUCT_UPDATED', 'PRODUCTS', product.ProductId, '', JSON.stringify(product), 'Product updated');
      return successResponse_({ ProductId: product.ProductId }, 'Product updated.');
    }

    var newId = generateId_('PRD');
    appendRow_('PRODUCTS', {
      ProductId: newId, Name: safeString_(product.Name), Description: safeString_(product.Description),
      Category: safeString_(product.Category) || PRODUCT_CATEGORIES[0], ImageUrl: safeString_(product.ImageUrl),
      Images: safeString_(product.Images), Price: price, Unit: safeString_(product.Unit) || 'per unit', EstWeight: safeNumber_(product.EstWeight),
      MinQty: safeNumber_(product.MinQty, 1), MaxQty: safeNumber_(product.MaxQty, 10),
      Available: safeBoolean_(product.Available, true), Featured: safeBoolean_(product.Featured, false), Active: true,
      BreedType: safeString_(product.BreedType) || 'Desi / Country', BirdType: safeString_(product.BirdType) || 'Hen',
      Source: safeString_(product.Source), AgeRange: safeString_(product.AgeRange),
      TypicalWeight: safeString_(product.TypicalWeight), FarmingMethod: safeString_(product.FarmingMethod),
      CreatedAt: nowIso_()
    });
    writeAudit_(currentEmail_(), currentRole_(), 'PRODUCT_CREATED', 'PRODUCTS', newId, '', JSON.stringify(product), 'Product created');
    return successResponse_({ ProductId: newId }, 'Product created.');
  });
}

function adminDeleteProduct(productId) {
  return safeInvoke_('adminDeleteProduct', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var existing = getRowById_('PRODUCTS', 'ProductId', productId);
    if (!existing) return errorResponse_('Product not found.', 'NOT_FOUND');
    updateRowFields_('PRODUCTS', 'ProductId', productId, { Active: false, Available: false });
    writeAudit_(currentEmail_(), currentRole_(), 'PRODUCT_DEACTIVATED', 'PRODUCTS', productId, 'Active=true', 'Active=false', 'Product deactivated');
    return successResponse_({}, 'Product deactivated.');
  });
}

/* ============================================================================
 * SECTION 12 — CUSTOMER LOOKUP & PROFILE HELPERS
 * ========================================================================== */

function upsertCustomer_(name, mobile, email) {
  var cleanMobile = String(mobile).replace(/\D/g, '').trim();
  var customers = sheetToObjects_('CUSTOMERS');
  
  for (var i = 0; i < customers.length; i++) {
    var cMobile = String(customers[i].Mobile || '').replace(/\D/g, '').trim();
    if (cMobile === cleanMobile || cMobile.endsWith(cleanMobile) || cleanMobile.endsWith(cMobile)) {
      return customers[i].CustomerId;
    }
  }
  
  var id = generateId_('CUS');
  appendRow_('CUSTOMERS', { CustomerId: id, Name: name, Mobile: mobile, Email: email || '', CreatedAt: nowIso_(), Active: true });
  return id;
}

function upsertAddress_(customerId, addressPayload) {
  var addrId = generateId_('ADDR');
  appendRow_('CUSTOMER_ADDRESSES', {
    AddressId: addrId, CustomerId: customerId, Label: safeString_(addressPayload.label) || 'Home',
    AddressLine: safeString_(addressPayload.addressLine), Area: safeString_(addressPayload.area),
    PIN: safeString_(addressPayload.pin), Lat: safeNumber_(addressPayload.lat), Lng: safeNumber_(addressPayload.lng),
    IsDefault: true, CreatedAt: nowIso_()
  });
  return addrId;
}

function getMyProfile(mobile) {
  return safeInvoke_('getMyProfile', function () {
    if (!mobile) return successResponse_(null);
    var cleanSearchMobile = String(mobile).replace(/\D/g, '').trim();
    
    var customers = sheetToObjects_('CUSTOMERS');
    var customer = null;
    
    for (var i = 0; i < customers.length; i++) {
      var cMobile = String(customers[i].Mobile || '').replace(/\D/g, '').trim();
      if (cMobile === cleanSearchMobile || cMobile.endsWith(cleanSearchMobile) || cleanSearchMobile.endsWith(cMobile)) {
        customer = customers[i];
        break;
      }
    }

    if (!customer) return successResponse_(null);
    var addresses = sheetToObjects_('CUSTOMER_ADDRESSES').filter(function (a) { return a.CustomerId === customer.CustomerId; });
    var walletBalance = getWalletBalance_(customer.CustomerId);
    var loyalty = getRowById_('LOYALTY', 'CustomerId', customer.CustomerId) || { Points: 0 };
    var referral = sheetToObjects_('REFERRALS').filter(function (r) { return r.ReferrerCustomerId === customer.CustomerId && !r.RefereeCustomerId; })[0];
    var referralCode = referral ? referral.ReferralCode : ensureReferralCode_(customer.CustomerId);
    var wholesale = sheetToObjects_('WHOLESALE_ACCOUNTS').filter(function (w) { return w.CustomerId === customer.CustomerId && safeBoolean_(w.Active); })[0];
    var subscriptions = sheetToObjects_('SUBSCRIPTIONS').filter(function (s) { return s.CustomerId === customer.CustomerId; });
    return successResponse_({
      customer: customer, addresses: addresses, walletBalance: walletBalance,
      loyaltyPoints: safeNumber_(loyalty.Points), loyaltyTier: loyaltyTierFor_(safeNumber_(loyalty.Points)),
      referralCode: referralCode, wholesaleAccount: wholesale || null, subscriptions: subscriptions
    });
  });
}

/* ============================================================================
 * SECTION 13 — COUPON EVALUATION
 * ========================================================================== */

function evaluateCoupon_(code, subtotal) {
  if (!code) return { valid: false, discount: 0 };
  var coupons = sheetToObjects_('COUPONS');
  var coupon = coupons.filter(function (c) { return c.Code === code && safeBoolean_(c.Active); })[0];
  if (!coupon) return { valid: false, discount: 0, message: 'Invalid or inactive coupon code.' };
  var now = new Date();
  if (coupon.ValidFrom && new Date(coupon.ValidFrom) > now) return { valid: false, discount: 0, message: 'Coupon not yet active.' };
  if (coupon.ValidTo && new Date(coupon.ValidTo) < now) return { valid: false, discount: 0, message: 'Coupon has expired.' };
  if (safeNumber_(coupon.UsageLimit) > 0 && safeNumber_(coupon.UsedCount) >= safeNumber_(coupon.UsageLimit)) {
    return { valid: false, discount: 0, message: 'Coupon usage limit reached.' };
  }
  if (subtotal < safeNumber_(coupon.MinOrder)) return { valid: false, discount: 0, message: 'Order does not meet coupon minimum of ' + formatCurrency(coupon.MinOrder) + '.' };
  var discount = coupon.DiscountType === 'PERCENT' ? (subtotal * safeNumber_(coupon.Value) / 100) : safeNumber_(coupon.Value);
  discount = Math.min(discount, subtotal);
  return { valid: true, discount: Math.round(discount * 100) / 100, couponId: coupon.CouponId };
}

/* ============================================================================
 * SECTION 14 — ORDER TOTALS
 * ========================================================================== */

function calculateOrderTotals_(cartLines, couponCode, customerId, opts) {
  opts = opts || {};
  var products = sheetToObjects_('PRODUCTS');
  var productMap = {};
  products.forEach(function (p) { productMap[p.ProductId] = p; });

  var wholesaleDiscountPct = 0;
  if (customerId) {
    var wholesale = sheetToObjects_('WHOLESALE_ACCOUNTS').filter(function (w) { return w.CustomerId === customerId && safeBoolean_(w.Active); })[0];
    if (wholesale) wholesaleDiscountPct = safeNumber_(wholesale.DiscountPercent);
  }

  var resolvedLines = [];
  var subtotal = 0;
  cartLines.forEach(function (line) {
    var product = productMap[line.productId];
    if (!product || !safeBoolean_(product.Active) || !safeBoolean_(product.Available)) throw new Error('Product unavailable: ' + line.productId);
    var qty = Math.max(1, Math.floor(safeNumber_(line.qty, 1)));
    if (qty < safeNumber_(product.MinQty, 1) || qty > safeNumber_(product.MaxQty, 99)) {
      throw new Error('Quantity for ' + product.Name + ' must be between ' + product.MinQty + ' and ' + product.MaxQty + '.');
    }
    var price = safeNumber_(product.Price);
    if (wholesaleDiscountPct > 0) price = Math.round(price * (1 - wholesaleDiscountPct / 100) * 100) / 100;
    var lineTotal = Math.round(price * qty * 100) / 100;
    subtotal += lineTotal;
    resolvedLines.push({
      productId: product.ProductId, name: product.Name, qty: qty,
      estWeight: safeNumber_(product.EstWeight) * qty, price: price, lineTotal: lineTotal
    });
  });

  var deliveryFee = safeNumber_(getConfigValue_('deliveryFee', DEFAULT_CONFIG.deliveryFee));
  var freeThreshold = safeNumber_(getConfigValue_('freeDeliveryThreshold', DEFAULT_CONFIG.freeDeliveryThreshold));
  if (subtotal >= freeThreshold || wholesaleDiscountPct > 0) deliveryFee = 0;

  var couponResult = evaluateCoupon_(couponCode, subtotal);
  var couponDiscount = couponResult.valid ? couponResult.discount : 0;

  var giftCardResult = { valid: false, applied: 0 };
  if (opts.giftCardCode) giftCardResult = previewGiftCard_(opts.giftCardCode);

  var taxPercent = safeNumber_(getConfigValue_('taxPercent', 0));
  var taxable = subtotal - couponDiscount;
  var tax = Math.round(taxable * taxPercent) / 100;

  var runningTotal = Math.round((subtotal + deliveryFee + tax - couponDiscount) * 100) / 100;

  var giftCardApplied = 0;
  if (giftCardResult.valid) giftCardApplied = Math.min(giftCardResult.balance, runningTotal);
  runningTotal = Math.round((runningTotal - giftCardApplied) * 100) / 100;

  var walletBalance = customerId ? getWalletBalance_(customerId) : 0;
  var walletRequested = safeNumber_(opts.walletAmount);
  var walletApplied = Math.min(walletBalance, walletRequested, runningTotal);
  walletApplied = Math.round(walletApplied * 100) / 100;
  var total = Math.round((runningTotal - walletApplied) * 100) / 100;

  return {
    lines: resolvedLines, subtotal: Math.round(subtotal * 100) / 100, deliveryFee: deliveryFee,
    discount: couponDiscount, tax: tax, total: total, couponResult: couponResult,
    wholesaleDiscountPercent: wholesaleDiscountPct, giftCardApplied: giftCardApplied,
    giftCardCode: giftCardResult.valid ? opts.giftCardCode.toUpperCase() : '',
    walletApplied: walletApplied, walletBalance: walletBalance
  };
}

function calculateCartPreview(cartLines, couponCode, mobile, walletAmount, giftCardCode) {
  return safeInvoke_('calculateCartPreview', function () {
    if (!cartLines || cartLines.length === 0) return errorResponse_('Cart is empty.', 'EMPTY_CART');
    var customerId = null;
    if (mobile) {
      var customer = sheetToObjects_('CUSTOMERS').filter(function (c) { return c.Mobile === mobile; })[0];
      if (customer) customerId = customer.CustomerId;
    }
    return successResponse_(calculateOrderTotals_(cartLines, couponCode, customerId, { walletAmount: walletAmount, giftCardCode: giftCardCode }));
  });
}

/* ============================================================================
 * SECTION 15 — ORDER LIFECYCLE & AUTO-DELIVERY SYNC
 * ========================================================================== */

function transitionOrderStatus_(orderId, newStatus, description) {
  var order = getRowById_('ORDERS', 'OrderId', orderId);
  if (!order) throw new Error('Order not found: ' + orderId);
  var allowed = ORDER_TRANSITIONS[order.Status] || [];
  if (allowed.indexOf(newStatus) === -1) {
    throw new Error('Invalid status transition from ' + order.Status + ' to ' + newStatus);
  }
  updateRowFields_('ORDERS', 'OrderId', orderId, { Status: newStatus, UpdatedAt: nowIso_() });
  writeAudit_(currentEmail_(), currentRole_(), 'ORDER_STATUS_CHANGED', 'ORDERS', orderId, order.Status, newStatus, description || '');
  
  // AUTO-CREATE OR SYNC DELIVERY ROW FOR PACKED / READY / OUT_FOR_DELIVERY ORDERS
  if (['PACKED', 'READY_FOR_DELIVERY', 'OUT_FOR_DELIVERY'].indexOf(newStatus) !== -1) {
    ensureDeliveryRecordExists_(orderId, order.DeliverySlotId, newStatus);
  }
  
  return true;
}

/**
 * Ensures a matching Delivery record exists in the DELIVERIES sheet
 */
function ensureDeliveryRecordExists_(orderId, slotId, orderStatus) {
  try {
    var deliveries = sheetToObjects_('DELIVERIES');
    var existing = deliveries.filter(function (d) { return d.OrderId === orderId; })[0];
    if (!existing) {
      appendRow_('DELIVERIES', {
        DeliveryId: generateId_('DEL'),
        OrderId: orderId,
        Status: orderStatus === 'OUT_FOR_DELIVERY' ? 'OUT_FOR_DELIVERY' : 'PENDING',
        SlotId: slotId || '',
        AssignedTo: '',
        DeliveredAt: '',
        Notes: 'Auto-created on order status transition'
      });
    } else if (orderStatus === 'OUT_FOR_DELIVERY' && existing.Status !== 'DELIVERED') {
      updateRowFields_('DELIVERIES', 'DeliveryId', existing.DeliveryId, { Status: 'OUT_FOR_DELIVERY' });
    }
  } catch (e) {
    Logger.log('ensureDeliveryRecordExists_ warning: ' + e);
  }
}

function incrementCouponUsage_(couponIdOrCode) {
  if (!couponIdOrCode) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var coupons = sheetToObjects_('COUPONS');
    var coupon = coupons.filter(function (c) { 
      return String(c.CouponId) === String(couponIdOrCode) || 
             safeString_(c.Code).toUpperCase() === safeString_(couponIdOrCode).toUpperCase(); 
    })[0];
    
    if (coupon) {
      var currentUsed = safeNumber_(coupon.UsedCount, 0);
      updateRowFields_('COUPONS', 'CouponId', coupon.CouponId, { UsedCount: currentUsed + 1 });
    }
  } finally {
    lock.releaseLock();
  }
}

function submitOrder(payload) {
  return safeInvoke_('submitOrder', function () {
    if (!payload || !payload.cart || payload.cart.length === 0) return errorResponse_('Your cart is empty.', 'EMPTY_CART');
    if (!safeString_(payload.name) || !safeString_(payload.mobile)) return errorResponse_('Name and mobile number are required.', 'VALIDATION_ERROR');
    if (!safeString_(payload.addressLine) || !safeString_(payload.pin)) return errorResponse_('Delivery address is required.', 'VALIDATION_ERROR');
    if (!payload.slotId) return errorResponse_('Please choose a delivery slot.', 'VALIDATION_ERROR');

    var radiusCheck = isWithinDeliveryRadius_(payload.lat, payload.lng);
    if (payload.lat && payload.lng && !radiusCheck.withinRadius) {
      return errorResponse_(radiusCheck.reason, 'OUTSIDE_DELIVERY_ZONE');
    }

    var slot = getRowById_('DELIVERY_SLOTS', 'SlotId', payload.slotId);
    if (!slot || !safeBoolean_(slot.Active)) return errorResponse_('Selected delivery slot is not available.', 'INVALID_SLOT');
    if (safeNumber_(slot.Booked) >= safeNumber_(slot.Capacity)) return errorResponse_('Selected delivery slot is fully booked.', 'SLOT_FULL');

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var orderId, totals, customerId, isFirstOrder;
    try {
      customerId = upsertCustomer_(payload.name, payload.mobile, payload.email);
      var priorOrders = sheetToObjects_('ORDERS').filter(function (o) { return o.CustomerId === customerId; });
      isFirstOrder = priorOrders.length === 0;

      try {
        totals = calculateOrderTotals_(payload.cart, payload.couponCode, customerId, { walletAmount: payload.walletAmount, giftCardCode: payload.giftCardCode });
      } catch (calcErr) {
        lock.releaseLock();
        return errorResponse_(calcErr.message, 'VALIDATION_ERROR');
      }

      var minOrder = safeNumber_(getConfigValue_('minimumOrder', DEFAULT_CONFIG.minimumOrder));
      if (totals.subtotal < minOrder) { lock.releaseLock(); return errorResponse_('Minimum order value is ' + formatCurrency(minOrder) + '.', 'BELOW_MINIMUM'); }

      var addressId = upsertAddress_(customerId, { label: 'Delivery', addressLine: payload.addressLine, area: payload.area, pin: payload.pin, lat: payload.lat, lng: payload.lng });

      orderId = generateId_('DM');
      appendRow_('ORDERS', {
        OrderId: orderId, CustomerId: customerId, Status: 'DRAFT', PaymentStatus: 'PENDING',
        Subtotal: totals.subtotal, DeliveryFee: totals.deliveryFee, Discount: totals.discount, Total: totals.total,
        AddressId: addressId, DeliverySlotId: payload.slotId, SpecialInstruction: safeString_(payload.instructions),
        CreatedAt: nowIso_(), UpdatedAt: nowIso_()
      });
      totals.lines.forEach(function (line) {
        appendRow_('ORDER_LINES', {
          LineId: generateId_('DM') + '-' + line.productId, OrderId: orderId, ProductId: line.productId,
          Qty: line.qty, EstWeight: line.estWeight, Price: line.price, LineTotal: line.lineTotal
        });
      });

      if (totals.couponResult && totals.couponResult.valid && totals.couponResult.couponId) {
        incrementCouponUsage_(totals.couponResult.couponId);
      }

      updateRowFields_('DELIVERY_SLOTS', 'SlotId', payload.slotId, { Booked: safeNumber_(slot.Booked) + 1 });

      if (totals.walletApplied > 0) debitWallet_(customerId, totals.walletApplied, 'Applied to order ' + orderId, orderId);
      if (totals.giftCardApplied > 0) consumeGiftCard_(totals.giftCardCode, totals.giftCardApplied, orderId);
      if (isFirstOrder && payload.referralCode) applyReferralCode_(payload.referralCode, customerId, orderId);

      writeAudit_(currentEmail_(), currentRole_(), 'ORDER_CREATED', 'ORDERS', orderId, '', totals.total, 'Order created via storefront');
      transitionOrderStatus_(orderId, 'PENDING_PAYMENT', 'Awaiting payment');

      if (totals.total <= 0) {
        updateRowFields_('ORDERS', 'OrderId', orderId, { PaymentStatus: 'PAID' });
        transitionOrderStatus_(orderId, 'PAID', 'Fully covered by wallet/gift card');
        transitionOrderStatus_(orderId, 'CONFIRMED', 'Auto-confirmed (zero balance due)');
        reserveInventoryForOrder_(orderId);
        transitionOrderStatus_(orderId, 'RESERVED', 'Live inventory reserved');
      }
    } finally {
      lock.releaseLock();
    }

    notify_('ORDER_CREATED', payload.email, { orderId: orderId, total: totals.total, name: payload.name });

    return successResponse_({
      orderId: orderId, total: totals.total, upiId: getConfigValue_('upiId', DEFAULT_CONFIG.upiId),
      upiQrUrl: getConfigValue_('upiQrUrl', ''), mobile: payload.mobile, walletApplied: totals.walletApplied,
      giftCardApplied: totals.giftCardApplied, fullyPaid: totals.total <= 0
    }, totals.total <= 0 ? 'Order placed and fully paid via wallet/gift card.' : 'Order placed. Please complete payment.');
  });
}

function getMyOrders(mobile) {
  return safeInvoke_('getMyOrders', function () {
    if (!mobile) return errorResponse_('Mobile number required.', 'VALIDATION_ERROR');
    var cleanSearchMobile = String(mobile).replace(/\D/g, '').trim();
    
    var customers = sheetToObjects_('CUSTOMERS');
    var matchingCustomerIds = [];
    
    customers.forEach(function (c) {
      var cMobile = String(c.Mobile || '').replace(/\D/g, '').trim();
      if (cMobile === cleanSearchMobile || cMobile.endsWith(cleanSearchMobile) || cleanSearchMobile.endsWith(cMobile)) {
        matchingCustomerIds.push(c.CustomerId);
      }
    });

    if (matchingCustomerIds.length === 0) return successResponse_([]);

    var orders = sheetToObjects_('ORDERS').filter(function (o) { 
      return matchingCustomerIds.indexOf(o.CustomerId) !== -1; 
    });
    
    var lines = sheetToObjects_('ORDER_LINES');
    var invoices = sheetToObjects_('INVOICES');
    
    orders.forEach(function (o) {
      o.lines = lines.filter(function (l) { return l.OrderId === o.OrderId; });
      var inv = invoices.filter(function (i) { return i.OrderId === o.OrderId; })[0];
      o.invoiceUrl = inv ? inv.PdfUrl : '';
    });
    
    orders.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_(orders);
  });
}

function getAdminOrders(statusFilter) {
  return safeInvoke_('getAdminOrders', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PACKING', 'DELIVERY']);
    var orders = sheetToObjects_('ORDERS');
    var customers = sheetToObjects_('CUSTOMERS');
    var customerMap = {};
    customers.forEach(function (c) { customerMap[c.CustomerId] = c; });
    var lines = sheetToObjects_('ORDER_LINES');
    var invoices = sheetToObjects_('INVOICES');
    var invoiceMap = {};
    invoices.forEach(function (i) { invoiceMap[i.OrderId] = i.PdfUrl; });

    orders.forEach(function (o) {
      o.customer = customerMap[o.CustomerId] || {};
      o.lines = lines.filter(function (l) { return l.OrderId === o.OrderId; });
      o.invoiceUrl = invoiceMap[o.OrderId] || '';
    });
    if (statusFilter) orders = orders.filter(function (o) { return o.Status === statusFilter; });
    orders.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_(orders);
  });
}

function adminTransitionOrder(orderId, newStatus) {
  return safeInvoke_('adminTransitionOrder', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PACKING', 'DELIVERY']);
    transitionOrderStatus_(orderId, newStatus, 'Manual transition by ' + currentEmail_());
    
    if (newStatus === 'PACKED') {
      try { generateInvoicePdf_(orderId); } catch (e) { Logger.log('Auto invoice generation warning: ' + e); }
    }
    if (newStatus === 'DELIVERED') {
      var deliveries = sheetToObjects_('DELIVERIES').filter(function (d) { return d.OrderId === orderId; })[0];
      if (deliveries) {
        updateRowFields_('DELIVERIES', 'DeliveryId', deliveries.DeliveryId, { Status: 'DELIVERED', DeliveredAt: nowIso_() });
      }
    }
    if (newStatus === 'COMPLETED') {
      awardOrderCompletionBenefits_(orderId);
    }
    return successResponse_({}, 'Order updated to ' + newStatus + '.');
  });
}

function awardOrderCompletionBenefits_(orderId) {
  try {
    var order = getRowById_('ORDERS', 'OrderId', orderId);
    if (!order) return;
    earnLoyaltyPoints_(order.CustomerId, safeNumber_(order.Total), orderId);
    settleReferralReward_(order.CustomerId, orderId);
  } catch (e) {
    Logger.log('awardOrderCompletionBenefits_ failed for ' + orderId + ': ' + e);
  }
}

function adminCancelOrder(orderId, reason) {
  return safeInvoke_('adminCancelOrder', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var order = getRowById_('ORDERS', 'OrderId', orderId);
    if (!order) return errorResponse_('Order not found.', 'NOT_FOUND');
    if (['RESERVED', 'PROCESSING', 'PACKED', 'READY_FOR_DELIVERY'].indexOf(order.Status) !== -1) {
      unreserveInventoryForOrder_(orderId);
    }
    var walletDebits = sheetToObjects_('WALLET_TRANSACTIONS').filter(function (t) { return t.RefId === orderId && t.Type === 'DEBIT'; });
    walletDebits.forEach(function (t) {
      creditWallet_(order.CustomerId, safeNumber_(t.Amount), 'Refund — order ' + orderId + ' cancelled', orderId);
    });
    transitionOrderStatus_(orderId, 'CANCELLED', reason || 'Cancelled by admin');
    return successResponse_({}, 'Order cancelled.' + (walletDebits.length ? ' Wallet balance refunded.' : ''));
  });
}

/* ============================================================================
 * SECTION 16 — PAYMENTS
 * ========================================================================== */

function submitPaymentReference(orderId, upiRef, amount) {
  return safeInvoke_('submitPaymentReference', function () {
    var order = getRowById_('ORDERS', 'OrderId', orderId);
    if (!order) return errorResponse_('Order not found.', 'NOT_FOUND');
    var paymentId = generateId_('PAY');
    appendRow_('PAYMENTS', {
      PaymentId: paymentId, OrderId: orderId, Amount: safeNumber_(amount, order.Total), Method: 'UPI',
      UpiRef: safeString_(upiRef), Timestamp: nowIso_(), Status: 'PENDING', VerificationStatus: 'PENDING'
    });
    updateRowFields_('ORDERS', 'OrderId', orderId, { PaymentStatus: 'PENDING' });
    writeAudit_(currentEmail_(), currentRole_(), 'PAYMENT_SUBMITTED', 'PAYMENTS', paymentId, '', upiRef, 'Customer submitted UPI reference');
    notify_('PAYMENT_SUBMITTED_ADMIN', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), { orderId: orderId, upiRef: upiRef });
    return successResponse_({ paymentId: paymentId }, 'Payment reference submitted. We will verify shortly.');
  });
}

function getAdminPayments(statusFilter) {
  return safeInvoke_('getAdminPayments', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var payments = sheetToObjects_('PAYMENTS');
    if (statusFilter) payments = payments.filter(function (p) { return p.VerificationStatus === statusFilter; });
    payments.sort(function (a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
    return successResponse_(payments);
  });
}

function adminVerifyPayment(paymentId, approve) {
  return safeInvoke_('adminVerifyPayment', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var payment = getRowById_('PAYMENTS', 'PaymentId', paymentId);
    if (!payment) return errorResponse_('Payment not found.', 'NOT_FOUND');
    
    if (approve) {
      updateRowFields_('PAYMENTS', 'PaymentId', paymentId, { Status: 'PAID', VerificationStatus: 'MANUAL_VERIFIED' });
      updateRowFields_('ORDERS', 'OrderId', payment.OrderId, { PaymentStatus: 'PAID' });
      writeAudit_(currentEmail_(), currentRole_(), 'PAYMENT_VERIFIED', 'PAYMENTS', paymentId, 'PENDING', 'PAID', 'Payment manually verified');
      
      try {
        transitionOrderStatus_(payment.OrderId, 'PAID', 'Payment verified by ' + currentEmail_());
        transitionOrderStatus_(payment.OrderId, 'CONFIRMED', 'Auto-confirmed after payment');
        reserveInventoryForOrder_(payment.OrderId);
        transitionOrderStatus_(payment.OrderId, 'RESERVED', 'Live inventory reserved');
      } catch (invErr) {
        Logger.log('Warning: Inventory reservation skipped during payment verify for ' + payment.OrderId + ': ' + invErr);
      }
      
      var order = getRowById_('ORDERS', 'OrderId', payment.OrderId);
      var customer = getRowById_('CUSTOMERS', 'CustomerId', order.CustomerId);
      notify_('PAYMENT_CONFIRMED', customer ? customer.Email : '', { orderId: payment.OrderId });
    } else {
      updateRowFields_('PAYMENTS', 'PaymentId', paymentId, { Status: 'FAILED', VerificationStatus: 'MANUAL_VERIFIED' });
      writeAudit_(currentEmail_(), currentRole_(), 'PAYMENT_REJECTED', 'PAYMENTS', paymentId, 'PENDING', 'FAILED', 'Payment rejected by ' + currentEmail_());
    }
    return successResponse_({}, approve ? 'Payment verified and order confirmed.' : 'Payment marked as failed.');
  });
}

/* ============================================================================
 * SECTION 17 — LIVE INVENTORY & LEDGER
 * ========================================================================== */

function recalcLiveInventorySummary_() {
  var batches = sheetToObjects_('LIVE_BATCHES').filter(function (b) { return b.Status === 'ACTIVE'; });
  var totals = { Available: 0, Reserved: 0, Processing: 0, Processed: 0, Mortality: 0, TotalLiveWeight: 0 };
  batches.forEach(function (b) {
    totals.Available += safeNumber_(b.QtyAvailable);
    totals.Reserved += safeNumber_(b.QtyReserved);
    totals.Processing += safeNumber_(b.QtyProcessing);
    totals.Processed += safeNumber_(b.QtyProcessed);
    totals.Mortality += safeNumber_(b.Mortality);
    totals.TotalLiveWeight += safeNumber_(b.TotalLiveWeight);
  });
  var totalBirds = totals.Available + totals.Reserved + totals.Processing;
  updateRowFields_('LIVE_INVENTORY', 'Id', 'SUMMARY', {
    TotalLiveBirds: totalBirds, Available: totals.Available, Reserved: totals.Reserved,
    Processing: totals.Processing, Processed: totals.Processed, Mortality: totals.Mortality,
    TotalLiveWeight: Math.round(totals.TotalLiveWeight * 100) / 100, LastUpdated: nowIso_()
  });
  return totals;
}

function getLiveInventorySummary() {
  return safeInvoke_('getLiveInventorySummary', function () {
    var summary = getRowById_('LIVE_INVENTORY', 'Id', 'SUMMARY');
    var minInv = safeNumber_(getConfigValue_('minimumLiveInventory', DEFAULT_CONFIG.minimumLiveInventory));
    var maxInv = safeNumber_(getConfigValue_('maximumLiveInventory', DEFAULT_CONFIG.maximumLiveInventory));
    return successResponse_({ summary: summary, minimum: minInv, maximum: maxInv, lowStock: safeNumber_(summary ? summary.Available : 0) < minInv });
  });
}

function reserveInventoryForOrder_(orderId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var lines = sheetToObjects_('ORDER_LINES').filter(function (l) { return l.OrderId === orderId; });
    var totalBirdsNeeded = lines.reduce(function (sum, l) { return sum + safeNumber_(l.Qty); }, 0);
    var batches = sheetToObjects_('LIVE_BATCHES')
      .filter(function (b) { return b.Status === 'ACTIVE' && safeNumber_(b.QtyAvailable) > 0; })
      .sort(function (a, b) { return new Date(a.ArrivalDate) - new Date(b.ArrivalDate); });

    var remaining = totalBirdsNeeded;
    for (var i = 0; i < batches.length && remaining > 0; i++) {
      var batch = batches[i];
      var take = Math.min(remaining, safeNumber_(batch.QtyAvailable));
      if (take <= 0) continue;
      updateRowFields_('LIVE_BATCHES', 'BatchId', batch.BatchId, {
        QtyAvailable: safeNumber_(batch.QtyAvailable) - take,
        QtyReserved: safeNumber_(batch.QtyReserved) + take
      });
      remaining -= take;
    }
    if (remaining > 0) {
      throw new Error('Insufficient live inventory to fulfill order ' + orderId + '. Short by ' + remaining + ' bird(s).');
    }
    recalcLiveInventorySummary_();
    writeAudit_(currentEmail_(), currentRole_(), 'INVENTORY_ADJUSTED', 'LIVE_BATCHES', orderId, '', 'RESERVED:' + totalBirdsNeeded, 'Reserved live inventory for order');
  } finally {
    lock.releaseLock();
  }
}

function unreserveInventoryForOrder_(orderId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var lines = sheetToObjects_('ORDER_LINES').filter(function (l) { return l.OrderId === orderId; });
    var totalBirds = lines.reduce(function (sum, l) { return sum + safeNumber_(l.Qty); }, 0);
    var batches = sheetToObjects_('LIVE_BATCHES').filter(function (b) { return safeNumber_(b.QtyReserved) > 0; });
    var remaining = totalBirds;
    for (var i = 0; i < batches.length && remaining > 0; i++) {
      var batch = batches[i];
      var give = Math.min(remaining, safeNumber_(batch.QtyReserved));
      if (give <= 0) continue;
      updateRowFields_('LIVE_BATCHES', 'BatchId', batch.BatchId, {
        QtyAvailable: safeNumber_(batch.QtyAvailable) + give,
        QtyReserved: safeNumber_(batch.QtyReserved) - give
      });
      remaining -= give;
    }
    recalcLiveInventorySummary_();
    writeAudit_(currentEmail_(), currentRole_(), 'INVENTORY_ADJUSTED', 'LIVE_BATCHES', orderId, '', 'UNRESERVED:' + totalBirds, 'Released reserved inventory for cancelled order');
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================================
 * SECTION 18 — SUPPLIERS & PROCUREMENT
 * ========================================================================== */

function getAdminSuppliers() {
  return safeInvoke_('getAdminSuppliers', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    return successResponse_(sheetToObjects_('SUPPLIERS'));
  });
}

function adminSaveSupplier(supplier) {
  return safeInvoke_('adminSaveSupplier', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    if (!supplier || !safeString_(supplier.BusinessName)) return errorResponse_('Business name is required.', 'VALIDATION_ERROR');
    if (supplier.SupplierId) {
      updateRowFields_('SUPPLIERS', 'SupplierId', supplier.SupplierId, {
        BusinessName: safeString_(supplier.BusinessName), ContactPerson: safeString_(supplier.ContactPerson),
        Phone: safeString_(supplier.Phone), WhatsApp: safeString_(supplier.WhatsApp), Address: safeString_(supplier.Address),
        Location: safeString_(supplier.Location), BreedType: safeString_(supplier.BreedType),
        MinOrder: safeNumber_(supplier.MinOrder), Price: safeNumber_(supplier.Price), AvgWeight: safeNumber_(supplier.AvgWeight),
        Capacity: safeNumber_(supplier.Capacity), LeadTime: safeNumber_(supplier.LeadTime), Reliability: safeString_(supplier.Reliability),
        Active: safeBoolean_(supplier.Active, true), Notes: safeString_(supplier.Notes)
      });
      writeAudit_(currentEmail_(), currentRole_(), 'SUPPLIER_UPDATED', 'SUPPLIERS', supplier.SupplierId, '', '', 'Supplier updated');
      return successResponse_({ SupplierId: supplier.SupplierId }, 'Supplier updated.');
    }
    var newId = generateId_('SUP');
    appendRow_('SUPPLIERS', {
      SupplierId: newId, BusinessName: safeString_(supplier.BusinessName), ContactPerson: safeString_(supplier.ContactPerson),
      Phone: safeString_(supplier.Phone), WhatsApp: safeString_(supplier.WhatsApp), Address: safeString_(supplier.Address),
      Location: safeString_(supplier.Location), BreedType: safeString_(supplier.BreedType), MinOrder: safeNumber_(supplier.MinOrder),
      Price: safeNumber_(supplier.Price), AvgWeight: safeNumber_(supplier.AvgWeight), Capacity: safeNumber_(supplier.Capacity),
      LeadTime: safeNumber_(supplier.LeadTime), Reliability: safeString_(supplier.Reliability), Active: true,
      Notes: safeString_(supplier.Notes), CreatedAt: nowIso_()
    });
    writeAudit_(currentEmail_(), currentRole_(), 'SUPPLIER_CREATED', 'SUPPLIERS', newId, '', '', 'Supplier created');
    return successResponse_({ SupplierId: newId }, 'Supplier created.');
  });
}

function computeProcurementRecommendation_() {
  var summary = getRowById_('LIVE_INVENTORY', 'Id', 'SUMMARY');
  var minInv = safeNumber_(getConfigValue_('minimumLiveInventory', DEFAULT_CONFIG.minimumLiveInventory));
  var maxInv = safeNumber_(getConfigValue_('maximumLiveInventory', DEFAULT_CONFIG.maximumLiveInventory));
  var safety = safeNumber_(getConfigValue_('safetyStock', DEFAULT_CONFIG.safetyStock));
  var available = safeNumber_(summary ? summary.Available : 0);
  var recommended = Math.max(0, (maxInv - available) + safety);
  return { available: available, minimum: minInv, maximum: maxInv, safetyStock: safety, lowStock: available < minInv, recommendedPurchase: recommended };
}

function checkLowInventory() {
  var rec = computeProcurementRecommendation_();
  if (rec.lowStock) {
    var existingDraftPOs = sheetToObjects_('PURCHASE_ORDERS').filter(function (po) { return po.Status === 'DRAFT' && po.Notes === 'AUTO_RECOMMENDED'; });
    if (existingDraftPOs.length === 0) {
      var suppliers = sheetToObjects_('SUPPLIERS').filter(function (s) { return safeBoolean_(s.Active); });
      var preferredSupplier = suppliers[0];
      var newId = generateId_('PO');
      appendRow_('PURCHASE_ORDERS', {
        PoId: newId, SupplierId: preferredSupplier ? preferredSupplier.SupplierId : '', Date: nowIso_(),
        RequiredDate: nowIso_(), Product: 'Desi Hen (Live)', Quantity: rec.recommendedPurchase,
        TargetWeight: rec.recommendedPurchase * 1.3, AgreedPrice: preferredSupplier ? preferredSupplier.Price : 0,
        EstimatedTotal: rec.recommendedPurchase * (preferredSupplier ? preferredSupplier.Price : 0),
        Status: 'DRAFT', Notes: 'AUTO_RECOMMENDED', CreatedAt: nowIso_()
      });
      notify_('LOW_INVENTORY', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), rec);
      writeAudit_('SYSTEM', 'SYSTEM', 'PO_CREATED', 'PURCHASE_ORDERS', newId, '', '', 'Auto-recommended PO created due to low inventory');
    }
  }
  return rec;
}

function adminCheckLowInventory() {
  return safeInvoke_('adminCheckLowInventory', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    return successResponse_(checkLowInventory());
  });
}

function getAdminPurchaseOrders() {
  return safeInvoke_('getAdminPurchaseOrders', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    var pos = sheetToObjects_('PURCHASE_ORDERS');
    var suppliers = sheetToObjects_('SUPPLIERS');
    var supplierMap = {};
    suppliers.forEach(function (s) { supplierMap[s.SupplierId] = s.BusinessName; });
    pos.forEach(function (po) { po.supplierName = supplierMap[po.SupplierId] || 'Unknown'; });
    pos.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_(pos);
  });
}

function adminCreatePurchaseOrder(po) {
  return safeInvoke_('adminCreatePurchaseOrder', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    if (!po || !po.SupplierId || !safeNumber_(po.Quantity)) return errorResponse_('Supplier and quantity are required.', 'VALIDATION_ERROR');
    var newId = generateId_('PO');
    appendRow_('PURCHASE_ORDERS', {
      PoId: newId, SupplierId: po.SupplierId, Date: nowIso_(), RequiredDate: po.RequiredDate || nowIso_(),
      Product: safeString_(po.Product) || 'Desi Hen (Live)', Quantity: safeNumber_(po.Quantity),
      TargetWeight: safeNumber_(po.TargetWeight), AgreedPrice: safeNumber_(po.AgreedPrice),
      EstimatedTotal: safeNumber_(po.Quantity) * safeNumber_(po.AgreedPrice), Status: 'DRAFT',
      Notes: safeString_(po.Notes), CreatedAt: nowIso_()
    });
    writeAudit_(currentEmail_(), currentRole_(), 'PO_CREATED', 'PURCHASE_ORDERS', newId, '', '', 'Purchase order created manually');
    return successResponse_({ PoId: newId }, 'Purchase order created.');
  });
}

function adminUpdatePurchaseOrderStatus(poId, status) {
  return safeInvoke_('adminUpdatePurchaseOrderStatus', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    var valid = ['DRAFT', 'SENT', 'ACCEPTED', 'PARTIAL', 'RECEIVED', 'CANCELLED'];
    if (valid.indexOf(status) === -1) return errorResponse_('Invalid PO status.', 'VALIDATION_ERROR');
    var po = getRowById_('PURCHASE_ORDERS', 'PoId', poId);
    if (!po) return errorResponse_('Purchase order not found.', 'NOT_FOUND');
    updateRowFields_('PURCHASE_ORDERS', 'PoId', poId, { Status: status });
    writeAudit_(currentEmail_(), currentRole_(), 'PO_STATUS_CHANGED', 'PURCHASE_ORDERS', poId, po.Status, status, '');
    if (status === 'ACCEPTED') notify_('PURCHASE_ORDER_ACCEPTED', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), { poId: poId });
    return successResponse_({}, 'Purchase order updated.');
  });
}

/* ============================================================================
 * SECTION 19 — RECEIVING & QC
 * ========================================================================== */

function adminReceivePurchaseOrder(payload) {
  return safeInvoke_('adminReceivePurchaseOrder', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT', 'FARM']);
    var po = getRowById_('PURCHASE_ORDERS', 'PoId', payload.poId);
    if (!po) return errorResponse_('Purchase order not found.', 'NOT_FOUND');

    var qtyReceived = safeNumber_(payload.qtyReceived);
    var rejected = safeNumber_(payload.rejected);
    var mortality = safeNumber_(payload.mortality);
    var accepted = Math.max(0, qtyReceived - rejected - mortality);
    var avgWeight = safeNumber_(payload.avgLiveWeight);
    var totalWeight = Math.round(accepted * avgWeight * 100) / 100;

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var receivingId, batchId, qcId;
    try {
      receivingId = generateId_('RECV');
      appendRow_('RECEIVING', {
        ReceivingId: receivingId, PoId: payload.poId, QtyOrdered: po.Quantity, QtyReceived: qtyReceived,
        Rejected: rejected, Mortality: mortality, Accepted: accepted, AvgLiveWeight: avgWeight,
        TotalLiveWeight: totalWeight, ReceivingDate: nowIso_(), QcStatus: accepted > 0 ? 'PASSED' : 'REJECTED'
      });

      batchId = generateId_('BATCH');
      appendRow_('LIVE_BATCHES', {
        BatchId: batchId, SupplierId: po.SupplierId, ReceivingId: receivingId, BirdType: 'Hen',
        Sex: safeString_(payload.sex) || 'Mixed', Breed: safeString_(payload.breed) || 'Desi', QtyReceived: qtyReceived,
        QtyAvailable: accepted, QtyReserved: 0, QtyProcessing: 0, QtyProcessed: 0, Mortality: mortality,
        AvgLiveWeight: avgWeight, TotalLiveWeight: totalWeight, ArrivalDate: nowIso_(),
        ExpectedProcessingDate: payload.expectedProcessingDate || '', Status: 'ACTIVE'
      });

      qcId = generateId_('QC');
      appendRow_('QC_RECORDS', {
        QcId: qcId, BatchId: batchId, SupplierId: po.SupplierId, ReceivedQty: qtyReceived, Healthy: accepted,
        Rejected: rejected, Mortality: mortality, Condition: safeString_(payload.condition) || 'Good',
        WeightCheck: 'OK', Notes: safeString_(payload.notes), Inspector: currentEmail_(), Timestamp: nowIso_(),
        Status: accepted > 0 ? 'PASSED' : 'REJECTED'
      });

      if (mortality > 0) {
        appendRow_('MORTALITY_LOG', { Id: generateId_('MORT'), BatchId: batchId, Qty: mortality, Date: nowIso_(), Reason: 'In-transit / receiving', RecordedBy: currentEmail_() });
      }

      updateRowFields_('PURCHASE_ORDERS', 'PoId', payload.poId, { Status: 'RECEIVED' });
      recalcLiveInventorySummary_();
    } finally {
      lock.releaseLock();
    }
    writeAudit_(currentEmail_(), currentRole_(), 'RECEIVING_COMPLETED', 'RECEIVING', receivingId, '', accepted, 'Received against PO ' + payload.poId);
    return successResponse_({ receivingId: receivingId, batchId: batchId, qcId: qcId, accepted: accepted }, 'Receiving recorded and live batch created.');
  });
}

function getAdminReceivingHistory() {
  return safeInvoke_('getAdminReceivingHistory', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT', 'FARM']);
    var records = sheetToObjects_('RECEIVING');
    records.sort(function (a, b) { return new Date(b.ReceivingDate) - new Date(a.ReceivingDate); });
    return successResponse_(records);
  });
}

function getAdminLiveBatches() {
  return safeInvoke_('getAdminLiveBatches', function () {
    requireRole_(['ADMIN', 'MANAGER', 'FARM', 'PROCESSING']);
    var batches = sheetToObjects_('LIVE_BATCHES');
    batches.sort(function (a, b) { return new Date(b.ArrivalDate) - new Date(a.ArrivalDate); });
    return successResponse_(batches);
  });
}

/* ============================================================================
 * SECTION 20 — PROCESSING MODULE
 * ========================================================================== */

function adminCreateProcessingBatch(payload) {
  return safeInvoke_('adminCreateProcessingBatch', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCESSING']);
    var batch = getRowById_('LIVE_BATCHES', 'BatchId', payload.sourceLiveBatchId);
    if (!batch) return errorResponse_('Live batch not found.', 'NOT_FOUND');

    var birdCount = safeNumber_(payload.birdCount);
    if (birdCount <= 0 || birdCount > safeNumber_(batch.QtyReserved) + safeNumber_(batch.QtyAvailable)) {
      return errorResponse_('Requested bird count exceeds birds available in this batch.', 'VALIDATION_ERROR');
    }
    var totalLiveWeight = safeNumber_(payload.totalLiveWeight) || (birdCount * safeNumber_(batch.AvgLiveWeight));
    var rejected = safeNumber_(payload.rejected);
    var mortality = safeNumber_(payload.mortality);
    var birdsProcessed = Math.max(0, birdCount - rejected - mortality);
    var dressedWeight = safeNumber_(payload.dressedWeight);
    var saleableWeight = safeNumber_(payload.saleableWeight) || dressedWeight;
    var waste = Math.max(0, totalLiveWeight - dressedWeight);
    var yieldPct = totalLiveWeight > 0 ? Math.round((dressedWeight / totalLiveWeight) * 10000) / 100 : 0;
    var saleableYieldPct = totalLiveWeight > 0 ? Math.round((saleableWeight / totalLiveWeight) * 10000) / 100 : 0;

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var procId;
    try {
      procId = generateId_('PROC');
      appendRow_('PROCESSING_BATCHES', {
        ProcBatchId: procId, SourceLiveBatchId: payload.sourceLiveBatchId, BirdCount: birdCount,
        TotalLiveWeight: totalLiveWeight, BirdsProcessed: birdsProcessed, Rejected: rejected, Mortality: mortality,
        DressedWeight: dressedWeight, SaleableWeight: saleableWeight, Waste: Math.round(waste * 100) / 100,
        YieldPct: yieldPct, SaleableYieldPct: saleableYieldPct, Status: 'COMPLETED', Date: nowIso_()
      });

      var fromReserved = Math.min(birdCount, safeNumber_(batch.QtyReserved));
      var fromAvailable = birdCount - fromReserved;
      updateRowFields_('LIVE_BATCHES', 'BatchId', batch.BatchId, {
        QtyReserved: safeNumber_(batch.QtyReserved) - fromReserved,
        QtyAvailable: safeNumber_(batch.QtyAvailable) - fromAvailable,
        QtyProcessed: safeNumber_(batch.QtyProcessed) + birdsProcessed,
        Mortality: safeNumber_(batch.Mortality) + mortality
      });
      if (mortality > 0) {
        appendRow_('MORTALITY_LOG', { Id: generateId_('MORT'), BatchId: batch.BatchId, Qty: mortality, Date: nowIso_(), Reason: 'Processing stage', RecordedBy: currentEmail_() });
      }

      if (payload.productId && saleableWeight > 0) {
        appendRow_('DRESSED_INVENTORY', {
          Id: generateId_('DRESSED'), ProcBatchId: procId, ProductId: payload.productId,
          Cut: safeString_(payload.cut) || 'Whole', Weight: saleableWeight, AvailableWeight: saleableWeight,
          ReservedWeight: 0, SoldWeight: 0, Wastage: Math.round(waste * 100) / 100, StorageStatus: 'FRESH',
          CreatedAt: nowIso_()
        });
      }
      recalcLiveInventorySummary_();
    } finally {
      lock.releaseLock();
    }
    writeAudit_(currentEmail_(), currentRole_(), 'PROCESSING_COMPLETED', 'PROCESSING_BATCHES', procId, '', birdsProcessed, 'Processing batch completed');
    return successResponse_({ procBatchId: procId, yieldPct: yieldPct, saleableYieldPct: saleableYieldPct }, 'Processing batch recorded.');
  });
}

function getAdminProcessingBatches() {
  return safeInvoke_('getAdminProcessingBatches', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCESSING']);
    var batches = sheetToObjects_('PROCESSING_BATCHES');
    batches.sort(function (a, b) { return new Date(b.Date) - new Date(a.Date); });
    return successResponse_(batches);
  });
}

function getAdminDressedInventory() {
  return safeInvoke_('getAdminDressedInventory', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCESSING', 'PACKING']);
    return successResponse_(sheetToObjects_('DRESSED_INVENTORY'));
  });
}

/* ============================================================================
 * SECTION 21 — LIVE FARM / PROCESSING PUBLIC STATUS PAGES
 * ========================================================================== */

function getLiveFarmStatus() {
  return safeInvoke_('getLiveFarmStatus', function () {
    recalcLiveInventorySummary_();
    
    var summary = getRowById_('LIVE_INVENTORY', 'Id', 'SUMMARY') || { TotalLiveBirds: 0, Available: 0, Reserved: 0, Processing: 0, Processed: 0 };
    var activeBatches = sheetToObjects_('LIVE_BATCHES').filter(function (b) { return b.Status === 'ACTIVE'; });
    var latestBatch = activeBatches.sort(function (a, b) { 
      return new Date(b.ArrivalDate || 0).getTime() - new Date(a.ArrivalDate || 0).getTime(); 
    })[0];

    return successResponse_({
      summary: summary,
      latestBatch: latestBatch || null,
      youtubeId: getConfigValue_('youtubeFarmLiveId', ''),
      lastUpdated: summary.LastUpdated || nowIso_()
    });
  });
}

function adminAddSampleLiveBatch() {
  return safeInvoke_('adminAddSampleLiveBatch', function () {
    requireRole_(['ADMIN', 'MANAGER', 'FARM']);
    var suppliers = sheetToObjects_('SUPPLIERS');
    var supplier = suppliers[0];
    var supplierId = supplier ? supplier.SupplierId : 'SUP-000001';
    
    var batchId = generateId_('BATCH');
    var receivingId = generateId_('RECV');
    
    appendRow_('RECEIVING', {
      ReceivingId: receivingId, PoId: 'SAMPLE-PO', QtyOrdered: 50, QtyReceived: 50,
      Rejected: 0, Mortality: 0, Accepted: 50, AvgLiveWeight: 1.3,
      TotalLiveWeight: 65, ReceivingDate: nowIso_(), QcStatus: 'PASSED'
    });

    appendRow_('LIVE_BATCHES', {
      BatchId: batchId, SupplierId: supplierId, ReceivingId: receivingId, BirdType: 'Hen',
      Sex: 'Mixed', Breed: 'Desi', QtyReceived: 50, QtyAvailable: 50, QtyReserved: 0,
      QtyProcessing: 0, QtyProcessed: 0, Mortality: 0, AvgLiveWeight: 1.3,
      TotalLiveWeight: 65, ArrivalDate: nowIso_(), ExpectedProcessingDate: '', Status: 'ACTIVE'
    });

    recalcLiveInventorySummary_();
    return successResponse_({ batchId: batchId, count: 50 }, '50 live Desi birds added to inventory batch ' + batchId + '.');
  });
}

function getLiveProcessingStatus() {
  return safeInvoke_('getLiveProcessingStatus', function () {
    var batches = sheetToObjects_('PROCESSING_BATCHES');
    batches.sort(function (a, b) { return new Date(b.Date) - new Date(a.Date); });
    var today = Utilities.formatDate(new Date(), getConfigValue_('timezone', 'Asia/Kolkata'), 'yyyy-MM-dd');
    var todaysBatches = batches.filter(function (b) { return String(b.Date).indexOf(today) === 0; });
    return successResponse_({
      youtubeId: getConfigValue_('youtubeProcessingLiveId', ''),
      currentBatch: batches[0] || null,
      todaysBatches: todaysBatches,
      lastUpdated: nowIso_()
    });
  });
}

/* ============================================================================
 * SECTION 22 — DELIVERY MANAGEMENT (FIXED ADMIN QUEUE)
 * ========================================================================== */

function getDeliverySlots() {
  try {
    var ss = ss_();
    var sheet = ss.getSheetByName('DELIVERY_SLOTS');
    if (!sheet || sheet.getLastRow() < 2) {
      if (!sheet) sheet = ss.insertSheet('DELIVERY_SLOTS');
      var headers = ['SlotId', 'Name', 'StartTime', 'EndTime', 'Capacity', 'Booked', 'Active'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var defaultSlots = [
        ['SLOT-001', 'Morning', '08:00', '11:00', 25, 0, true],
        ['SLOT-002', 'Afternoon', '12:00', '15:00', 25, 0, true],
        ['SLOT-003', 'Evening', '17:00', '20:00', 25, 0, true]
      ];
      sheet.getRange(2, 1, defaultSlots.length, defaultSlots[0].length).setValues(defaultSlots);
    }
    var rows = sheetToObjects_('DELIVERY_SLOTS');
    var slots = rows.filter(function (s) { return safeBoolean_(s.Active, true); });
    slots.forEach(function (s) { 
      s.remaining = Math.max(0, Number(s.Capacity || 25) - Number(s.Booked || 0)); 
    });
    return successResponse_(slots);
  } catch (err) {
    var fallback = [
      { SlotId: 'SLOT-001', Name: 'Morning', StartTime: '08:00', EndTime: '11:00', remaining: 25 },
      { SlotId: 'SLOT-002', Name: 'Afternoon', StartTime: '12:00', EndTime: '15:00', remaining: 25 },
      { SlotId: 'SLOT-003', Name: 'Evening', StartTime: '17:00', EndTime: '20:00', remaining: 25 }
    ];
    return successResponse_(fallback);
  }
}

function adminAssignDelivery(orderId, assignedTo) {
  return safeInvoke_('adminAssignDelivery', function () {
    requireRole_(['ADMIN', 'MANAGER', 'DELIVERY']);
    var order = getRowById_('ORDERS', 'OrderId', orderId);
    if (!order) return errorResponse_('Order not found.', 'NOT_FOUND');
    
    var existing = sheetToObjects_('DELIVERIES').filter(function (d) { return d.OrderId === orderId; })[0];
    if (existing) {
      updateRowFields_('DELIVERIES', 'DeliveryId', existing.DeliveryId, { AssignedTo: assignedTo, Status: 'ASSIGNED' });
    } else {
      appendRow_('DELIVERIES', {
        DeliveryId: generateId_('DEL'), OrderId: orderId, Status: 'ASSIGNED', SlotId: order.DeliverySlotId || '',
        AssignedTo: assignedTo, DeliveredAt: '', Notes: ''
      });
    }
    transitionOrderStatus_(orderId, 'OUT_FOR_DELIVERY', 'Assigned to ' + assignedTo);
    writeAudit_(currentEmail_(), currentRole_(), 'DELIVERY_ASSIGNED', 'DELIVERIES', orderId, '', assignedTo, '');
    return successResponse_({}, 'Delivery assigned to ' + assignedTo + '.');
  });
}

function getAdminDeliveries() {
  return safeInvoke_('getAdminDeliveries', function () {
    requireRole_(['ADMIN', 'MANAGER', 'DELIVERY']);
    
    var deliveries = sheetToObjects_('DELIVERIES');
    var orders = sheetToObjects_('ORDERS');
    var customers = sheetToObjects_('CUSTOMERS');
    var addresses = sheetToObjects_('CUSTOMER_ADDRESSES');
    var slots = sheetToObjects_('DELIVERY_SLOTS');
    var orderLines = sheetToObjects_('ORDER_LINES');
    var products = sheetToObjects_('PRODUCTS');

    var customerMap = {}, addressMap = {}, slotMap = {}, productMap = {}, orderMap = {};
    customers.forEach(function (c) { customerMap[c.CustomerId] = c; });
    addresses.forEach(function (a) { addressMap[a.AddressId] = a; });
    slots.forEach(function (s) { slotMap[s.SlotId] = s.Name + ' (' + s.StartTime + '-' + s.EndTime + ')'; });
    products.forEach(function (p) { productMap[p.ProductId] = p.Name; });
    orders.forEach(function (o) { orderMap[o.OrderId] = o; });

    // Sync packed / ready orders into deliveries list if missing
    orders.forEach(function (o) {
      if (['PACKED', 'READY_FOR_DELIVERY', 'OUT_FOR_DELIVERY', 'DELIVERED'].indexOf(o.Status) !== -1) {
        var dExists = deliveries.filter(function (d) { return d.OrderId === o.OrderId; })[0];
        if (!dExists) {
          ensureDeliveryRecordExists_(o.OrderId, o.DeliverySlotId, o.Status);
        }
      }
    });

    deliveries = sheetToObjects_('DELIVERIES'); // Reload enriched list

    deliveries.forEach(function (d) {
      var o = orderMap[d.OrderId] || {};
      var c = customerMap[o.CustomerId] || {};
      var a = addressMap[o.AddressId] || {};
      var lines = orderLines.filter(function (l) { return l.OrderId === d.OrderId; });
      
      d.customerName = c.Name || 'Customer';
      d.customerMobile = c.Mobile || '';
      d.addressLine = (a.AddressLine || '') + (a.Area ? ', ' + a.Area : '') + (a.PIN ? ' - ' + a.PIN : '');
      d.slotName = slotMap[d.SlotId || o.DeliverySlotId] || 'Standard Slot';
      d.orderTotal = o.Total || 0;
      d.orderStatus = o.Status || 'UNKNOWN';
      d.itemsSummary = lines.map(function (l) { return (productMap[l.ProductId] || l.ProductId) + ' x' + l.Qty; }).join(', ');
    });

    deliveries.sort(function (a, b) { return new Date(b._row || 0) - new Date(a._row || 0); });
    return successResponse_(deliveries);
  });
}

function updateDeliveryStatus() {
  var deliveries = sheetToObjects_('DELIVERIES').filter(function (d) { return d.Status === 'OUT_FOR_DELIVERY' || d.Status === 'ASSIGNED'; });
  return { checked: deliveries.length };
}

/* ============================================================================
 * SECTION 23 — INVOICE & DOC GENERATION (WITH BRAND LOGO)
 * ========================================================================== */

function generateInvoicePdf_(orderId) {
  var order = getRowById_('ORDERS', 'OrderId', orderId);
  if (!order) throw new Error('Order record not found: ' + orderId);

  var existingInvoice = sheetToObjects_('INVOICES').filter(function (i) { return i.OrderId === orderId; })[0];
  if (existingInvoice) return existingInvoice;

  var customer = getRowById_('CUSTOMERS', 'CustomerId', order.CustomerId) || {};
  var address = getRowById_('CUSTOMER_ADDRESSES', 'AddressId', order.AddressId) || {};
  var lines = sheetToObjects_('ORDER_LINES').filter(function (l) { return l.OrderId === orderId; });
  var products = sheetToObjects_('PRODUCTS');
  var productMap = {};
  products.forEach(function (p) { productMap[p.ProductId] = p; });

  var invoiceId = generateId_('INV');
  var brandName = getConfigValue_('brandName', DEFAULT_CONFIG.brandName);
  var tagline = getConfigValue_('tagline', DEFAULT_CONFIG.tagline);
  var logoUrl = getConfigValue_('logoUrl', '');

  var doc = DocumentApp.create(invoiceId + '_' + orderId);
  var body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

  // EMBED OFFICIAL BRAND LOGO IF CONFIGURED
  if (logoUrl) {
    try {
      var fileId = extractDriveId_(logoUrl);
      if (fileId) {
        var logoBlob = DriveApp.getFileById(fileId).getBlob();
        var logoPara = body.appendParagraph('');
        logoPara.appendInlineImage(logoBlob).setWidth(100).setHeight(50);
      }
    } catch (e) {
      Logger.log('Invoice logo embed warning: ' + e);
    }
  }

  var title = body.appendParagraph(brandName);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(tagline).setItalic(true);
  body.appendParagraph(getConfigValue_('businessAddress', DEFAULT_CONFIG.businessAddress));
  body.appendHorizontalRule();

  var meta = body.appendTable([
    ['Invoice #', invoiceId],
    ['Order #', orderId],
    ['Date', Utilities.formatDate(new Date(), getConfigValue_('timezone', 'Asia/Kolkata'), 'dd MMM yyyy')],
    ['Payment Status', safeString_(order.PaymentStatus || 'PENDING')]
  ]);
  meta.setBorderWidth(0);

  body.appendParagraph('Bill To').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(customer.Name || 'Valued Customer');
  body.appendParagraph(customer.Mobile || '');
  if (address.AddressLine) {
    body.appendParagraph(address.AddressLine + ', ' + (address.Area || '') + ' - ' + (address.PIN || ''));
  }

  body.appendParagraph('Order Items').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var tableData = [['Item', 'Qty', 'Est. Weight (kg)', 'Rate', 'Amount']];
  lines.forEach(function (line) {
    var product = productMap[line.ProductId];
    tableData.push([
      product ? product.Name : line.ProductId, 
      String(line.Qty || 1), 
      String(line.EstWeight || 0),
      formatCurrency(line.Price || 0), 
      formatCurrency(line.LineTotal || 0)
    ]);
  });
  
  if (tableData.length === 1) {
    tableData.push(['Desi Hen Order Item', '1', '1.0', formatCurrency(order.Total), formatCurrency(order.Total)]);
  }

  var table = body.appendTable(tableData);
  table.getRow(0).editAsText().setBold(true);

  body.appendParagraph(' ');
  var totalsTable = body.appendTable([
    ['Subtotal', formatCurrency(order.Subtotal || order.Total)],
    ['Delivery Fee', formatCurrency(order.DeliveryFee || 0)],
    ['Discount', '-' + formatCurrency(order.Discount || 0)],
    ['Total', formatCurrency(order.Total || 0)]
  ]);
  totalsTable.getRow(3).editAsText().setBold(true);

  body.appendHorizontalRule();
  body.appendParagraph('Support: ' + getConfigValue_('supportPhone', DEFAULT_CONFIG.supportPhone) + ' | ' + getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail));
  body.appendParagraph('Thank you for choosing ' + brandName + '. ' + tagline).setItalic(true);

  doc.saveAndClose();

  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  var folderId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.INVOICE_FOLDER_ID);
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var pdfFile = folder.createFile(pdfBlob).setName(invoiceId + '.pdf');
  docFile.setTrashed(true);

  var pdfUrl = pdfFile.getUrl();
  appendRow_('INVOICES', { InvoiceId: invoiceId, OrderId: orderId, PdfUrl: pdfUrl, CreatedAt: nowIso_() });
  writeAudit_(currentEmail_(), currentRole_(), 'INVOICE_GENERATED', 'INVOICES', invoiceId, '', pdfUrl, 'Invoice PDF generated for order ' + orderId);
  
  return { InvoiceId: invoiceId, OrderId: orderId, PdfUrl: pdfUrl };
}

function adminGenerateInvoice(orderId) {
  return safeInvoke_('adminGenerateInvoice', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PACKING']);
    var invoice = generateInvoicePdf_(orderId);
    if (!invoice) throw new Error('Failed to generate PDF invoice.');
    return successResponse_(invoice, 'Invoice PDF generated successfully.');
  });
}

/* ============================================================================
 * SECTION 24 — REVIEWS
 * ========================================================================== */

function submitReview(payload) {
  return safeInvoke_('submitReview', function () {
    var order = getRowById_('ORDERS', 'OrderId', payload.orderId);
    if (!order || ['DELIVERED', 'COMPLETED'].indexOf(order.Status) === -1) {
      return errorResponse_('Reviews can only be submitted for delivered or completed orders.', 'VALIDATION_ERROR');
    }
    var rating = Math.min(5, Math.max(1, safeNumber_(payload.rating, 5)));
    var reviewId = generateId_('REVIEW');
    appendRow_('REVIEWS', {
      ReviewId: reviewId, CustomerId: order.CustomerId, OrderId: payload.orderId, ProductId: payload.productId || '',
      Rating: rating, Comment: safeString_(payload.comment), PhotoUrl: safeString_(payload.photoUrl),
      CreatedAt: nowIso_(), VerifiedPurchase: true, AdminApproved: false, Published: false
    });
    writeAudit_(currentEmail_(), currentRole_(), 'REVIEW_SUBMITTED', 'REVIEWS', reviewId, '', rating, 'Customer review submitted');
    return successResponse_({ reviewId: reviewId }, 'Thank you for your review! It will appear after moderation.');
  });
}

function getPublishedReviews() {
  return safeInvoke_('getPublishedReviews', function () {
    var reviews = sheetToObjects_('REVIEWS').filter(function (r) { return safeBoolean_(r.Published); });
    var customers = sheetToObjects_('CUSTOMERS');
    var customerMap = {};
    
    customers.forEach(function (c) { 
      customerMap[c.CustomerId] = c; 
    });

    // Attach customer's real name and profile photo to each review
    reviews.forEach(function (r) {
      var c = customerMap[r.CustomerId] || {};
      r.customerName = c.Name || 'Verified Customer';
      r.customerPhoto = c.PhotoUrl || '';
    });

    reviews.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_(reviews.slice(0, 20));
  });
}

function getAdminReviews() {
  return safeInvoke_('getAdminReviews', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var reviews = sheetToObjects_('REVIEWS');
    reviews.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_(reviews);
  });
}

function adminModerateReview(reviewId, approve) {
  return safeInvoke_('adminModerateReview', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    updateRowFields_('REVIEWS', 'ReviewId', reviewId, { AdminApproved: !!approve, Published: !!approve });
    writeAudit_(currentEmail_(), currentRole_(), 'REVIEW_MODERATED', 'REVIEWS', reviewId, '', approve, '');
    return successResponse_({}, approve ? 'Review published.' : 'Review rejected.');
  });
}

/* ============================================================================
 * SECTION 25 — NOTIFICATIONS
 * ========================================================================== */

function notify_(event, recipient, payload) {
  var notifId = generateId_('NOTIF');
  var status = 'SKIPPED';
  try {
    if (recipient && recipient.indexOf('@') > -1) {
      var brandName = getConfigValue_('brandName', DEFAULT_CONFIG.brandName);
      var subject = brandName + ' — ' + event.replace(/_/g, ' ');
      var body = buildNotificationBody_(event, payload);
      MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: body });
      status = 'SENT';
    }
  } catch (e) {
    status = 'FAILED';
    Logger.log('Notification failed: ' + e);
  }
  appendRow_('NOTIFICATIONS', {
    NotifId: notifId, Event: event, Recipient: recipient || '', Channel: 'EMAIL',
    Payload: JSON.stringify(payload || {}), Status: status, CreatedAt: nowIso_()
  });
  return status;
}

function buildNotificationBody_(event, payload) {
  payload = payload || {};
  var brandName = esc_(getConfigValue_('brandName', DEFAULT_CONFIG.brandName));
  var tagline = esc_(getConfigValue_('tagline', DEFAULT_CONFIG.tagline));
  var locationName = esc_(getConfigValue_('locationName', DEFAULT_CONFIG.locationName));
  var supportPhone = esc_(getConfigValue_('supportPhone', DEFAULT_CONFIG.supportPhone));
  var supportEmail = esc_(getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail));
  var logoUrl = getConfigValue_('logoUrl', '');
  var directLogoUrl = logoUrl ? ('https://lh3.googleusercontent.com/d/' + extractDriveId_(logoUrl)) : '';
  var webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) {}

  // Context titles & action cards based on event
  var headerTitle = 'Notification Update';
  var statusBadgeText = 'UPDATE';
  var statusBadgeColor = '#7a4a2b';
  var ctaText = 'Visit DesiMurga™';
  var ctaUrl = webAppUrl || '#';
  var contentBodyHtml = '';

  if (event === 'ORDER_CREATED') {
    headerTitle = 'Order Received — Awaiting Payment';
    statusBadgeText = 'ORDER PLACED';
    statusBadgeColor = '#c98a2c';
    ctaText = 'Track Order & Complete Payment';
    ctaUrl = webAppUrl ? (webAppUrl + '#my-orders') : '#';
    contentBodyHtml = 
      '<p style="font-size:15px;color:#2b211a;margin:0 0 14px;line-height:1.5;">Hi <strong>' + esc_(payload.name || 'Valued Customer') + '</strong>,</p>' +
      '<p style="font-size:14px;color:#5c4a3a;margin:0 0 16px;line-height:1.6;">Thank you for ordering with DesiMurga™! Your fresh Desi chicken order has been recorded and live inventory is reserved.</p>' +
      '<div style="background:#fdf8f2;border:1.5px solid #f0e6d8;border-radius:14px;padding:16px;margin-bottom:18px;">' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
          '<tr><td style="color:#8a7a6d;padding:4px 0;">Order Reference:</td><td style="text-align:right;font-weight:bold;color:#5c3620;">' + esc_(payload.orderId) + '</td></tr>' +
          '<tr><td style="color:#8a7a6d;padding:4px 0;">Total Payable Amount:</td><td style="text-align:right;font-weight:bold;font-size:16px;color:#7a4a2b;">' + formatCurrency(payload.total) + '</td></tr>' +
          '<tr><td style="color:#8a7a6d;padding:4px 0;">Preparation:</td><td style="text-align:right;font-weight:600;color:#3f7d4c;">Dressed Fresh to Order</td></tr>' +
        '</table>' +
      '</div>' +
      '<p style="font-size:13px;color:#5c4a3a;margin:0 0 8px;line-height:1.5;">Please complete payment via UPI to dispatch your order on schedule.</p>';
  } 
  else if (event === 'PAYMENT_CONFIRMED') {
    headerTitle = 'Payment Confirmed & Order Reserved';
    statusBadgeText = 'CONFIRMED';
    statusBadgeColor = '#3f7d4c';
    ctaText = 'View Live Farm & Tracking';
    ctaUrl = webAppUrl ? (webAppUrl + '#my-orders') : '#';
    contentBodyHtml = 
      '<p style="font-size:14px;color:#5c4a3a;margin:0 0 16px;line-height:1.6;">We have verified your UPI payment for Order <strong>' + esc_(payload.orderId) + '</strong>. Your Desi hen is assigned for same-day dressing.</p>';
  }
  else if (event === 'ORDER_PACKED') {
    headerTitle = 'Freshly Packed & Quality Checked';
    statusBadgeText = 'PACKED';
    statusBadgeColor = '#3f7d4c';
    ctaText = 'Track Dispatch Status';
    contentBodyHtml = '<p style="font-size:14px;color:#5c4a3a;line-height:1.6;">Your chicken cut for order <strong>' + esc_(payload.orderId) + '</strong> has been dressed, vacuum sealed, and chilled for immediate slot dispatch.</p>';
  }
  else if (event === 'OUT_FOR_DELIVERY') {
    headerTitle = 'Out for Slot Delivery';
    statusBadgeText = 'ON THE WAY';
    statusBadgeColor = '#c98a2c';
    contentBodyHtml = '<p style="font-size:14px;color:#5c4a3a;line-height:1.6;">Our delivery rider is on the way with your insulated fresh package within our Garia service radius.</p>';
  }
  else if (event === 'ORDER_DELIVERED') {
    headerTitle = 'Delivered Fresh to Your Door';
    statusBadgeText = 'DELIVERED';
    statusBadgeColor = '#3f7d4c';
    ctaText = 'Rate & Review Your Order';
    ctaUrl = webAppUrl ? (webAppUrl + '#my-orders') : '#';
    contentBodyHtml = '<p style="font-size:14px;color:#5c4a3a;line-height:1.6;">Your order <strong>' + esc_(payload.orderId) + '</strong> has been delivered. Enjoy authentic Sunday Desi chicken with your family!</p>';
  }
  else {
    contentBodyHtml = '<p style="font-size:14px;color:#5c4a3a;line-height:1.6;">' + esc_(JSON.stringify(payload)) + '</p>';
  }

  // Master Email Template
  return '<!DOCTYPE html>' +
  '<html>' +
  '<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>' +
  '<body style="margin:0;padding:0;background-color:#f5ece0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">' +
    '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#f5ece0;padding:24px 12px;">' +
      '<tr><td align="center">' +
        '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width:540px;background-color:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 6px 24px rgba(90,60,30,0.08);border:1px solid #ebdccb;">' +
          
          // Header Banner with Logo
          '<tr><td style="background:linear-gradient(135deg, #7a4a2b 0%, #5c3620 100%);padding:28px 24px;text-align:center;">' +
            (directLogoUrl ? '<img src="' + esc_(directLogoUrl) + '" alt="DesiMurga" width="60" height="60" style="border-radius:50%;border:2px solid rgba(255,255,255,0.3);margin-bottom:10px;object-fit:cover;">' : '<div style="font-size:36px;margin-bottom:6px;">🐓</div>') +
            '<h1 style="color:#ffffff;font-size:24px;font-weight:900;margin:0;letter-spacing:-0.5px;">' + brandName + '</h1>' +
            '<p style="color:#f5ece0;font-size:12px;margin:4px 0 0;opacity:0.85;font-weight:500;">' + tagline + '</p>' +
          '</td></tr>' +

          // Main Card Content
          '<tr><td style="padding:28px 24px;">' +
            '<div style="display:inline-block;padding:4px 12px;background:' + statusBadgeColor + '15;color:' + statusBadgeColor + ';border-radius:20px;font-size:11px;font-weight:800;letter-spacing:0.5px;margin-bottom:12px;">' + statusBadgeText + '</div>' +
            '<h2 style="color:#5c3620;font-size:18px;font-weight:800;margin:0 0 16px;line-height:1.3;">' + headerTitle + '</h2>' +
            contentBodyHtml +
            
            // CTA Button
            '<div style="text-align:center;margin:24px 0 10px;">' +
              '<a href="' + ctaUrl + '" target="_blank" style="display:inline-block;background:#7a4a2b;color:#ffffff;text-decoration:none;font-size:14px;font-weight:800;padding:12px 28px;border-radius:12px;box-shadow:0 3px 10px rgba(122,74,43,0.25);">' + ctaText + ' &rarr;</a>' +
            '</div>' +
          '</td></tr>' +

          // Freshness Guarantee Strip
          '<tr><td style="background:#fdf8f2;border-top:1px solid #f0e6d8;padding:16px 24px;text-align:center;">' +
            '<table width="100%" border="0" cellspacing="0" cellpadding="0">' +
              '<tr>' +
                '<td align="center" style="font-size:11px;color:#7a4a2b;font-weight:700;">' +
                  '🌱 100% Free-Range Desi &nbsp;•&nbsp; 🔪 Dressed Same-Day &nbsp;•&nbsp; ❄️ Zero Cold Storage' +
                '</td>' +
              '</tr>' +
            '</table>' +
          '</td></tr>' +

          // Footer
          '<tr><td style="background:#2b211a;color:#d9c7b3;padding:24px;text-align:center;font-size:11px;line-height:1.6;">' +
            '<p style="margin:0 0 6px;color:#ffffff;font-weight:700;">DesiMurga™ Kolkata Delivery Hub</p>' +
            '<p style="margin:0 0 10px;color:#a89888;">📍 ' + locationName + ' (5 km Service Area)</p>' +
            '<p style="margin:0;color:#a89888;">Need help? Call/WhatsApp: <strong style="color:#ffffff;">' + supportPhone + '</strong> | <a href="mailto:' + supportEmail + '" style="color:#d9c7b3;text-decoration:underline;">' + supportEmail + '</a></p>' +
          '</td></tr>' +

        '</table>' +
      '</td></tr>' +
    '</table>' +
  '</body></html>';
}
/* ============================================================================
 * SECTION 26 — COUPONS (ADMIN)
 * ========================================================================== */

function adminSaveCoupon(coupon) {
  return safeInvoke_('adminSaveCoupon', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    if (!coupon || !safeString_(coupon.Code)) return errorResponse_('Coupon code is required.', 'VALIDATION_ERROR');
    var newId = generateId_('COUPON');
    appendRow_('COUPONS', {
      CouponId: newId, Code: safeString_(coupon.Code).toUpperCase(), DiscountType: coupon.DiscountType || 'FLAT',
      Value: safeNumber_(coupon.Value), MinOrder: safeNumber_(coupon.MinOrder), ValidFrom: coupon.ValidFrom || '',
      ValidTo: coupon.ValidTo || '', Active: true, UsageLimit: safeNumber_(coupon.UsageLimit), UsedCount: 0
    });
    writeAudit_(currentEmail_(), currentRole_(), 'COUPON_CREATED', 'COUPONS', newId, '', '', 'Coupon created');
    return successResponse_({ CouponId: newId }, 'Coupon created.');
  });
}

function getAdminCoupons() {
  return safeInvoke_('getAdminCoupons', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(sheetToObjects_('COUPONS'));
  });
}

/* ============================================================================
 * SECTION 27 — EXPENSES
 * ========================================================================== */

function adminSaveExpense(expense) {
  return safeInvoke_('adminSaveExpense', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var newId = generateId_('EXP');
    appendRow_('EXPENSES', { ExpenseId: newId, Category: safeString_(expense.Category), Amount: safeNumber_(expense.Amount), Date: expense.Date || nowIso_(), Notes: safeString_(expense.Notes) });
    writeAudit_(currentEmail_(), currentRole_(), 'EXPENSE_RECORDED', 'EXPENSES', newId, '', expense.Amount, '');
    return successResponse_({ ExpenseId: newId }, 'Expense recorded.');
  });
}

function getAdminExpenses() {
  return safeInvoke_('getAdminExpenses', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(sheetToObjects_('EXPENSES'));
  });
}

/* ============================================================================
 * SECTION 28 — ADMIN DASHBOARD & REPORTING
 * ========================================================================== */

function getAdminDashboardData() {
  return safeInvoke_('getAdminDashboardData', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT', 'FARM', 'PROCESSING', 'PACKING', 'DELIVERY']);
    var orders = sheetToObjects_('ORDERS');
    var timezone = getConfigValue_('timezone', 'Asia/Kolkata');
    var today = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
    var todaysOrders = orders.filter(function (o) { return String(o.CreatedAt).indexOf(today) === 0; });

    var todayStats = {
      sales: todaysOrders.filter(function (o) { return o.PaymentStatus === 'PAID'; }).reduce(function (s, o) { return s + safeNumber_(o.Total); }, 0),
      orders: todaysOrders.length,
      paid: todaysOrders.filter(function (o) { return o.PaymentStatus === 'PAID'; }).length,
      pending: todaysOrders.filter(function (o) { return o.PaymentStatus === 'PENDING'; }).length,
      cancelled: todaysOrders.filter(function (o) { return o.Status === 'CANCELLED'; }).length
    };
    todayStats.avgOrderValue = todayStats.orders > 0 ? Math.round((todaysOrders.reduce(function (s, o) { return s + safeNumber_(o.Total); }, 0) / todayStats.orders) * 100) / 100 : 0;

    var liveInv = getRowById_('LIVE_INVENTORY', 'Id', 'SUMMARY') || { TotalLiveBirds: 0, Available: 0, Reserved: 0, Processing: 0 };
    var procBatches = sheetToObjects_('PROCESSING_BATCHES');
    var todaysProc = procBatches.filter(function (p) { return String(p.Date).indexOf(today) === 0; });

    var deliveries = sheetToObjects_('DELIVERIES');
    var deliveryStats = {
      pending: orders.filter(function (o) { return o.Status === 'READY_FOR_DELIVERY'; }).length,
      packed: orders.filter(function (o) { return o.Status === 'PACKED'; }).length,
      outForDelivery: deliveries.filter(function (d) { return d.Status === 'OUT_FOR_DELIVERY' || d.Status === 'ASSIGNED'; }).length,
      delivered: deliveries.filter(function (d) { return d.Status === 'DELIVERED'; }).length
    };

    var procurement = computeProcurementRecommendation_();
    var pendingPOs = sheetToObjects_('PURCHASE_ORDERS').filter(function (p) { return ['DRAFT', 'SENT'].indexOf(p.Status) !== -1; }).length;

    var openTickets = sheetToObjects_('SUPPORT_TICKETS').filter(function (t) { return t.Status === 'OPEN' || t.Status === 'IN_PROGRESS'; }).length;
    var pendingReturns = sheetToObjects_('RETURNS').filter(function (r) { return r.Status === 'REQUESTED'; }).length;
    var activeSubscriptions = sheetToObjects_('SUBSCRIPTIONS').filter(function (s) { return s.Status === 'ACTIVE'; }).length;
    var ap = getAccountsPayableSummary_();

    return successResponse_({
      today: todayStats,
      liveFarm: liveInv,
      processing: { todayCount: todaysProc.length, completed: todaysProc.filter(function (p) { return p.Status === 'COMPLETED'; }).length, avgYield: todaysProc.length ? Math.round(todaysProc.reduce(function (s, p) { return s + safeNumber_(p.YieldPct); }, 0) / todaysProc.length * 100) / 100 : 0 },
      delivery: deliveryStats,
      procurement: { lowStockAlert: procurement.lowStock, recommendedPurchase: procurement.recommendedPurchase, pendingPOs: pendingPOs },
      growth: { openTickets: openTickets, pendingReturns: pendingReturns, activeSubscriptions: activeSubscriptions, accountsPayableOutstanding: ap.outstanding }
    });
  });
}

function generateDailySummary() {
  var orders = sheetToObjects_('ORDERS');
  var timezone = getConfigValue_('timezone', 'Asia/Kolkata');
  
  if (orders.length === 0) return { message: 'No orders found.' };

  var byDate = {};
  orders.forEach(function (o) {
    if (!o.CreatedAt) return;
    var dateStr = '';
    try {
      dateStr = Utilities.formatDate(new Date(o.CreatedAt), timezone, 'yyyy-MM-dd');
    } catch (e) {
      dateStr = String(o.CreatedAt).substring(0, 10);
    }
    
    if (!byDate[dateStr]) {
      byDate[dateStr] = { sales: 0, orders: 0, paid: 0, pending: 0, cancelled: 0 };
    }
    
    var stat = byDate[dateStr];
    stat.orders++;
    
    if (o.PaymentStatus === 'PAID') {
      stat.sales += safeNumber_(o.Total);
      stat.paid++;
    } else if (o.PaymentStatus === 'PENDING') {
      stat.pending++;
    }
    
    if (o.Status === 'CANCELLED') {
      stat.cancelled++;
    }
  });

  Object.keys(byDate).forEach(function (dateStr) {
    var stat = byDate[dateStr];
    var avgOrderValue = stat.orders > 0 ? Math.round((stat.sales / stat.orders) * 100) / 100 : 0;
    var rowData = {
      Date: dateStr,
      Sales: Math.round(stat.sales * 100) / 100,
      Orders: stat.orders,
      Paid: stat.paid,
      Pending: stat.pending,
      Cancelled: stat.cancelled,
      AvgOrderValue: avgOrderValue
    };

    var existingRowIndex = findRowIndexById_('DAILY_SUMMARY', 'Date', dateStr);
    if (existingRowIndex === -1) {
      appendRow_('DAILY_SUMMARY', rowData);
    } else {
      updateRowFields_('DAILY_SUMMARY', 'Date', dateStr, rowData);
    }
  });

  SpreadsheetApp.flush();
  return { datesProcessed: Object.keys(byDate).length };
}

function getSalesReport() {
  return safeInvoke_('getSalesReport', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    generateDailySummary();
    
    var summaries = sheetToObjects_('DAILY_SUMMARY');
    summaries.sort(function (a, b) { 
      return String(b.Date || '').localeCompare(String(a.Date || '')); 
    });
    return successResponse_(summaries);
  });
}

function getProfitabilityReport() {
  return safeInvoke_('getProfitabilityReport', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var orders = sheetToObjects_('ORDERS').filter(function (o) { return o.PaymentStatus === 'PAID'; });
    var sales = orders.reduce(function (s, o) { return s + safeNumber_(o.Total); }, 0);
    var expenses = sheetToObjects_('EXPENSES').reduce(function (s, e) { return s + safeNumber_(e.Amount); }, 0);
    var pos = sheetToObjects_('PURCHASE_ORDERS').filter(function (p) { return p.Status === 'RECEIVED'; });
    var procurementCost = pos.reduce(function (s, p) { return s + safeNumber_(p.EstimatedTotal); }, 0);
    var contribution = Math.round((sales - procurementCost - expenses) * 100) / 100;
    return successResponse_({ sales: Math.round(sales * 100) / 100, procurementCost: Math.round(procurementCost * 100) / 100, expenses: Math.round(expenses * 100) / 100, contribution: contribution });
  });
}

/* ============================================================================
 * SECTION 29 — USER MANAGEMENT
 * ========================================================================== */

function getAdminUsers() {
  return safeInvoke_('getAdminUsers', function () {
    requireRole_(['ADMIN']);
    return successResponse_(sheetToObjects_('USERS'));
  });
}

function adminSaveUser(user) {
  return safeInvoke_('adminSaveUser', function () {
    requireRole_(['ADMIN']);
    if (!user || !safeString_(user.Email) || ROLE_LIST.indexOf(user.Role) === -1) return errorResponse_('Valid email and role are required.', 'VALIDATION_ERROR');
    var existing = sheetToObjects_('USERS').filter(function (u) { return u.Email === user.Email; })[0];
    if (existing) {
      updateRowFields_('USERS', 'UserId', existing.UserId, { Name: safeString_(user.Name), Role: user.Role, Active: user.Active !== false });
    } else {
      appendRow_('USERS', { UserId: generateId_('USER'), Email: safeString_(user.Email), Name: safeString_(user.Name), Role: user.Role, Active: true, CreatedAt: nowIso_() });
    }
    CacheService.getScriptCache().remove('DM_ROLE_' + user.Email);
    writeAudit_(currentEmail_(), currentRole_(), 'USER_SAVED', 'USERS', user.Email, '', user.Role, '');
    return successResponse_({}, 'User saved.');
  });
}

/* ============================================================================
 * SECTION 30 — AUDIT LOG VIEWER
 * ========================================================================== */

function getAdminAuditLog() {
  return safeInvoke_('getAdminAuditLog', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var logs = sheetToObjects_('AUDIT_LOG');
    logs.sort(function (a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
    return successResponse_(logs.slice(0, 300));
  });
}

/* ============================================================================
 * SECTION 31A — WALLET / STORE CREDIT SERVICE
 * ========================================================================== */

function getWalletBalance_(customerId) {
  var txns = sheetToObjects_('WALLET_TRANSACTIONS').filter(function (t) { return t.CustomerId === customerId; });
  if (txns.length === 0) return 0;
  return safeNumber_(txns[txns.length - 1].Balance);
}

function creditWallet_(customerId, amount, reason, refId) {
  if (amount <= 0) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var balance = getWalletBalance_(customerId) + amount;
    appendRow_('WALLET_TRANSACTIONS', {
      TxnId: generateId_('WALLET'), CustomerId: customerId, Type: 'CREDIT', Amount: amount,
      Reason: reason, RefId: refId || '', Balance: Math.round(balance * 100) / 100, CreatedAt: nowIso_()
    });
  } finally { lock.releaseLock(); }
  writeAudit_(currentEmail_(), currentRole_(), 'WALLET_CREDITED', 'WALLET_TRANSACTIONS', customerId, '', amount, reason);
}

function debitWallet_(customerId, amount, reason, refId) {
  if (amount <= 0) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var currentBalance = getWalletBalance_(customerId);
    if (amount > currentBalance) throw new Error('Insufficient wallet balance.');
    var balance = currentBalance - amount;
    appendRow_('WALLET_TRANSACTIONS', {
      TxnId: generateId_('WALLET'), CustomerId: customerId, Type: 'DEBIT', Amount: amount,
      Reason: reason, RefId: refId || '', Balance: Math.round(balance * 100) / 100, CreatedAt: nowIso_()
    });
  } finally { lock.releaseLock(); }
  writeAudit_(currentEmail_(), currentRole_(), 'WALLET_DEBITED', 'WALLET_TRANSACTIONS', customerId, '', amount, reason);
}

function getMyWallet(mobile) {
  return safeInvoke_('getMyWallet', function () {
    var customer = sheetToObjects_('CUSTOMERS').filter(function (c) { return c.Mobile === mobile; })[0];
    if (!customer) return successResponse_({ balance: 0, transactions: [] });
    var txns = sheetToObjects_('WALLET_TRANSACTIONS').filter(function (t) { return t.CustomerId === customer.CustomerId; });
    txns.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_({ balance: getWalletBalance_(customer.CustomerId), transactions: txns });
  });
}

function adminAdjustWallet(mobile, amount, reason) {
  return safeInvoke_('adminAdjustWallet', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var customer = sheetToObjects_('CUSTOMERS').filter(function (c) { return c.Mobile === mobile; })[0];
    if (!customer) return errorResponse_('Customer not found for this mobile number.', 'NOT_FOUND');
    var amt = safeNumber_(amount);
    if (amt > 0) creditWallet_(customer.CustomerId, amt, reason || 'Manual admin credit', '');
    else debitWallet_(customer.CustomerId, Math.abs(amt), reason || 'Manual admin debit', '');
    return successResponse_({ balance: getWalletBalance_(customer.CustomerId) }, 'Wallet adjusted.');
  });
}

/* ============================================================================
 * SECTION 31B — LOYALTY PROGRAM
 * ========================================================================== */

function loyaltyTierFor_(points) {
  if (points >= 2000) return 'Gold';
  if (points >= 600) return 'Silver';
  return 'Bronze';
}

function earnLoyaltyPoints_(customerId, orderTotal, orderId) {
  var earnRate = safeNumber_(getConfigValue_('loyaltyEarnRatePercent', DEFAULT_CONFIG.loyaltyEarnRatePercent));
  var pointsEarned = Math.floor(orderTotal * earnRate / 100);
  if (pointsEarned <= 0) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var row = getRowById_('LOYALTY', 'CustomerId', customerId);
    if (row) updateRowFields_('LOYALTY', 'CustomerId', customerId, { Points: safeNumber_(row.Points) + pointsEarned, UpdatedAt: nowIso_() });
    else appendRow_('LOYALTY', { Id: generateId_('LOYALTY'), CustomerId: customerId, Points: pointsEarned, UpdatedAt: nowIso_() });
  } finally { lock.releaseLock(); }
  writeAudit_('SYSTEM', 'SYSTEM', 'LOYALTY_EARNED', 'LOYALTY', customerId, '', pointsEarned, 'Earned on completed order ' + orderId);
}

function redeemLoyaltyPoints(mobile, pointsToRedeem) {
  return safeInvoke_('redeemLoyaltyPoints', function () {
    var customer = sheetToObjects_('CUSTOMERS').filter(function (c) { return c.Mobile === mobile; })[0];
    if (!customer) return errorResponse_('Customer not found.', 'NOT_FOUND');
    var loyalty = getRowById_('LOYALTY', 'CustomerId', customer.CustomerId);
    var currentPoints = loyalty ? safeNumber_(loyalty.Points) : 0;
    var minRedeem = safeNumber_(getConfigValue_('loyaltyMinRedeem', DEFAULT_CONFIG.loyaltyMinRedeem));
    var pts = safeNumber_(pointsToRedeem);
    if (pts < minRedeem) return errorResponse_('Minimum ' + minRedeem + ' points required to redeem.', 'VALIDATION_ERROR');
    if (pts > currentPoints) return errorResponse_('You only have ' + currentPoints + ' points available.', 'VALIDATION_ERROR');
    var pointValue = safeNumber_(getConfigValue_('loyaltyPointValue', DEFAULT_CONFIG.loyaltyPointValue));
    var cashValue = Math.round(pts * pointValue * 100) / 100;
    updateRowFields_('LOYALTY', 'CustomerId', customer.CustomerId, { Points: currentPoints - pts, UpdatedAt: nowIso_() });
    creditWallet_(customer.CustomerId, cashValue, 'Redeemed ' + pts + ' loyalty points', '');
    writeAudit_(currentEmail_(), currentRole_(), 'LOYALTY_REDEEMED', 'LOYALTY', customer.CustomerId, currentPoints, currentPoints - pts, 'Redeemed for wallet credit');
    return successResponse_({ walletCredit: cashValue, remainingPoints: currentPoints - pts }, 'Redeemed ' + pts + ' points for ' + formatCurrency(cashValue) + ' wallet credit.');
  });
}

/* ============================================================================
 * SECTION 31C — REFERRAL PROGRAM
 * ========================================================================== */

function ensureReferralCode_(customerId) {
  var existing = sheetToObjects_('REFERRALS').filter(function (r) { return r.ReferrerCustomerId === customerId && !r.RefereeCustomerId; })[0];
  if (existing) return existing.ReferralCode;
  var customer = getRowById_('CUSTOMERS', 'CustomerId', customerId);
  var code = (customer ? safeString_(customer.Name).replace(/[^A-Za-z]/g, '').substring(0, 5).toUpperCase() : 'DESI') + Math.floor(1000 + Math.random() * 9000);
  appendRow_('REFERRALS', { ReferralId: generateId_('REF'), ReferrerCustomerId: customerId, ReferralCode: code, RefereeCustomerId: '', Status: 'ACTIVE', RewardAmount: 0, CreatedAt: nowIso_() });
  return code;
}

function applyReferralCode_(code, refereeCustomerId, orderId) {
  if (!code) return;
  var referral = sheetToObjects_('REFERRALS').filter(function (r) { return r.ReferralCode === safeString_(code).toUpperCase() && !r.RefereeCustomerId; })[0];
  if (!referral) return;
  if (referral.ReferrerCustomerId === refereeCustomerId) return;
  updateRowFields_('REFERRALS', 'ReferralId', referral.ReferralId, { RefereeCustomerId: refereeCustomerId, Status: 'PENDING' });
  var refereeDiscount = safeNumber_(getConfigValue_('refereeDiscountAmount', DEFAULT_CONFIG.refereeDiscountAmount));
  if (refereeDiscount > 0) creditWallet_(refereeCustomerId, refereeDiscount, 'Welcome credit — referral signup', orderId);
  writeAudit_('SYSTEM', 'SYSTEM', 'REFERRAL_APPLIED', 'REFERRALS', referral.ReferralId, '', refereeCustomerId, 'Referral code applied at checkout');
}

function settleReferralReward_(refereeCustomerId, orderId) {
  var referral = sheetToObjects_('REFERRALS').filter(function (r) { return r.RefereeCustomerId === refereeCustomerId && r.Status === 'PENDING'; })[0];
  if (!referral) return;
  var reward = safeNumber_(getConfigValue_('referralRewardAmount', DEFAULT_CONFIG.referralRewardAmount));
  creditWallet_(referral.ReferrerCustomerId, reward, 'Referral reward — friend completed first order', orderId);
  updateRowFields_('REFERRALS', 'ReferralId', referral.ReferralId, { Status: 'COMPLETED', RewardAmount: reward });
  writeAudit_('SYSTEM', 'SYSTEM', 'REFERRAL_REWARDED', 'REFERRALS', referral.ReferralId, 'PENDING', 'COMPLETED', 'Reward credited to referrer wallet');
}

function getMyReferralCode(mobile) {
  return safeInvoke_('getMyReferralCode', function () {
    if (!mobile) return errorResponse_('Mobile number required.', 'VALIDATION_ERROR');
    
    var cleanSearchMobile = String(mobile).replace(/\D/g, '').trim();
    var customers = sheetToObjects_('CUSTOMERS');
    var customer = null;
    
    for (var i = 0; i < customers.length; i++) {
      var cMobile = String(customers[i].Mobile || '').replace(/\D/g, '').trim();
      if (cMobile === cleanSearchMobile || cMobile.endsWith(cleanSearchMobile) || cleanSearchMobile.endsWith(cMobile)) {
        customer = customers[i];
        break;
      }
    }
    
    if (!customer) {
      return errorResponse_('Place an order first to get your referral code.', 'NOT_FOUND');
    }
    
    var code = ensureReferralCode_(customer.CustomerId);
    var referrals = sheetToObjects_('REFERRALS').filter(function (r) { 
      return r.ReferrerCustomerId === customer.CustomerId; 
    });
    
    return successResponse_({ code: code, referrals: referrals });
  });
}

function getAdminReferrals() {
  return safeInvoke_('getAdminReferrals', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(sheetToObjects_('REFERRALS'));
  });
}

/* ============================================================================
 * SECTION 31D — GIFT CARDS
 * ========================================================================== */

function adminIssueGiftCard(value, issuedTo, expiresAt) {
  return safeInvoke_('adminIssueGiftCard', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var amount = safeNumber_(value);
    if (amount <= 0) return errorResponse_('Gift card value must be greater than zero.', 'VALIDATION_ERROR');
    var code = 'DM-GIFT-' + Utilities.getUuid().split('-')[0].toUpperCase();
    var cardId = generateId_('GIFT');
    appendRow_('GIFT_CARDS', { CardId: cardId, Code: code, InitialValue: amount, Balance: amount, IssuedTo: safeString_(issuedTo), ExpiresAt: expiresAt || '', Active: true, CreatedAt: nowIso_() });
    writeAudit_(currentEmail_(), currentRole_(), 'GIFT_CARD_ISSUED', 'GIFT_CARDS', cardId, '', amount, 'Issued to ' + issuedTo);
    return successResponse_({ code: code, cardId: cardId }, 'Gift card issued: ' + code);
  });
}

function previewGiftCard_(code) {
  var card = sheetToObjects_('GIFT_CARDS').filter(function (g) { return g.Code === safeString_(code).toUpperCase() && safeBoolean_(g.Active); })[0];
  if (!card) return { valid: false, balance: 0 };
  if (card.ExpiresAt && new Date(card.ExpiresAt) < new Date()) return { valid: false, balance: 0, message: 'Gift card has expired.' };
  if (safeNumber_(card.Balance) <= 0) return { valid: false, balance: 0, message: 'Gift card has no remaining balance.' };
  return { valid: true, balance: safeNumber_(card.Balance) };
}

function consumeGiftCard_(code, amount, orderId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var card = sheetToObjects_('GIFT_CARDS').filter(function (g) { return g.Code === safeString_(code).toUpperCase(); })[0];
    if (!card) return;
    var newBalance = Math.max(0, safeNumber_(card.Balance) - amount);
    updateRowFields_('GIFT_CARDS', 'CardId', card.CardId, { Balance: newBalance, Active: newBalance > 0 });
  } finally { lock.releaseLock(); }
  writeAudit_(currentEmail_(), currentRole_(), 'GIFT_CARD_REDEEMED', 'GIFT_CARDS', code, '', amount, 'Applied to order ' + orderId);
}

function checkGiftCardBalance(code) {
  return safeInvoke_('checkGiftCardBalance', function () { return successResponse_(previewGiftCard_(code)); });
}

function getAdminGiftCards() {
  return safeInvoke_('getAdminGiftCards', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var cards = sheetToObjects_('GIFT_CARDS');
    
    cards.sort(function (a, b) { 
      var dB = new Date(b.CreatedAt || 0).getTime();
      var dA = new Date(a.CreatedAt || 0).getTime();
      return (isNaN(dB) ? 0 : dB) - (isNaN(dA) ? 0 : dA); 
    });

    return successResponse_(cards);
  });
}

/* ============================================================================
 * SECTION 31E — WHOLESALE / B2B ACCOUNTS
 * ========================================================================== */

function adminCreateWholesaleAccount(payload) {
  return safeInvoke_('adminCreateWholesaleAccount', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    if (!payload || !safeString_(payload.mobile) || !safeString_(payload.businessName)) return errorResponse_('Mobile and business name are required.', 'VALIDATION_ERROR');
    var customerId = upsertCustomer_(payload.name || payload.businessName, payload.mobile, payload.email);
    var accountId = generateId_('WHS');
    appendRow_('WHOLESALE_ACCOUNTS', {
      AccountId: accountId, CustomerId: customerId, BusinessName: safeString_(payload.businessName), GSTIN: safeString_(payload.gstin),
      CreditLimit: safeNumber_(payload.creditLimit), CreditUsed: 0, PaymentTerms: safeString_(payload.paymentTerms) || 'Prepaid',
      DiscountPercent: safeNumber_(payload.discountPercent, safeNumber_(getConfigValue_('wholesaleDefaultDiscountPercent', DEFAULT_CONFIG.wholesaleDefaultDiscountPercent))),
      Active: true, CreatedAt: nowIso_()
    });
    writeAudit_(currentEmail_(), currentRole_(), 'WHOLESALE_ACCOUNT_CREATED', 'WHOLESALE_ACCOUNTS', accountId, '', '', 'B2B account created for ' + payload.businessName);
    return successResponse_({ accountId: accountId }, 'Wholesale account created.');
  });
}

function getAdminWholesaleAccounts() {
  return safeInvoke_('getAdminWholesaleAccounts', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var accounts = sheetToObjects_('WHOLESALE_ACCOUNTS');
    var customers = sheetToObjects_('CUSTOMERS');
    var customerMap = {};
    customers.forEach(function (c) { customerMap[c.CustomerId] = c; });
    accounts.forEach(function (a) { a.customer = customerMap[a.CustomerId] || {}; });
    return successResponse_(accounts);
  });
}

/* ============================================================================
 * SECTION 31F — SUBSCRIPTIONS & RECURRING ORDERS
 * ========================================================================== */

function nextOrderDateFor_(frequency, fromDate) {
  var d = new Date(fromDate || new Date());
  if (frequency === 'WEEKLY') d.setDate(d.getDate() + 7);
  else if (frequency === 'BIWEEKLY') d.setDate(d.getDate() + 14);
  else d.setMonth(d.getMonth() + 1);
  return d.toISOString();
}

function createSubscription(payload) {
  return safeInvoke_('createSubscription', function () {
    if (!payload || !payload.mobile || !payload.productId || !payload.slotId) return errorResponse_('Mobile, product, and delivery slot are required.', 'VALIDATION_ERROR');
    if (SUBSCRIPTION_FREQUENCIES.indexOf(payload.frequency) === -1) return errorResponse_('Invalid frequency.', 'VALIDATION_ERROR');
    var customerId = upsertCustomer_(payload.name, payload.mobile, payload.email);
    var addressId = upsertAddress_(customerId, { label: 'Subscription', addressLine: payload.addressLine, area: payload.area, pin: payload.pin, lat: payload.lat, lng: payload.lng });
    var subId = generateId_('SUB');
    appendRow_('SUBSCRIPTIONS', {
      SubscriptionId: subId, CustomerId: customerId, ProductId: payload.productId, Qty: safeNumber_(payload.qty, 1),
      Frequency: payload.frequency, NextOrderDate: nextOrderDateFor_(payload.frequency), AddressId: addressId,
      SlotId: payload.slotId, Status: 'ACTIVE', CreatedAt: nowIso_(), LastOrderId: ''
    });
    writeAudit_(currentEmail_(), currentRole_(), 'SUBSCRIPTION_CREATED', 'SUBSCRIPTIONS', subId, '', '', 'Recurring order created');
    return successResponse_({ subscriptionId: subId }, 'Subscription created. Your first recurring order will be placed on schedule.');
  });
}

function getMySubscriptions(mobile) {
  return safeInvoke_('getMySubscriptions', function () {
    if (!mobile) return successResponse_([]);
    var cleanSearchMobile = String(mobile).replace(/\D/g, '').trim();
    
    var customers = sheetToObjects_('CUSTOMERS');
    var customer = null;
    
    for (var i = 0; i < customers.length; i++) {
      var cMobile = String(customers[i].Mobile || '').replace(/\D/g, '').trim();
      if (cMobile === cleanSearchMobile || cMobile.endsWith(cleanSearchMobile) || cleanSearchMobile.endsWith(cMobile)) {
        customer = customers[i];
        break;
      }
    }
    
    if (!customer) return successResponse_([]);
    return successResponse_(sheetToObjects_('SUBSCRIPTIONS').filter(function (s) { 
      return s.CustomerId === customer.CustomerId; 
    }));
  });
}

function getAdminSubscriptions() {
  return safeInvoke_('getAdminSubscriptions', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var subs = sheetToObjects_('SUBSCRIPTIONS');
    var customers = sheetToObjects_('CUSTOMERS');
    var products = sheetToObjects_('PRODUCTS');
    var customerMap = {}, productMap = {};
    customers.forEach(function (c) { customerMap[c.CustomerId] = c; });
    products.forEach(function (p) { productMap[p.ProductId] = p; });
    subs.forEach(function (s) { s.customer = customerMap[s.CustomerId] || {}; s.productName = (productMap[s.ProductId] || {}).Name || s.ProductId; });
    subs.sort(function (a, b) { return new Date(a.NextOrderDate) - new Date(b.NextOrderDate); });
    return successResponse_(subs);
  });
}

function updateSubscriptionStatus(subscriptionId, status) {
  return safeInvoke_('updateSubscriptionStatus', function () {
    if (SUBSCRIPTION_STATUSES.indexOf(status) === -1) return errorResponse_('Invalid subscription status.', 'VALIDATION_ERROR');
    updateRowFields_('SUBSCRIPTIONS', 'SubscriptionId', subscriptionId, { Status: status });
    writeAudit_(currentEmail_(), currentRole_(), 'SUBSCRIPTION_STATUS_CHANGED', 'SUBSCRIPTIONS', subscriptionId, '', status, '');
    return successResponse_({}, 'Subscription ' + status.toLowerCase() + '.');
  });
}

function processDueSubscriptions() {
  var subs = sheetToObjects_('SUBSCRIPTIONS').filter(function (s) { return s.Status === 'ACTIVE' && new Date(s.NextOrderDate) <= new Date(); });
  var placed = 0;
  subs.forEach(function (sub) {
    try {
      var customer = getRowById_('CUSTOMERS', 'CustomerId', sub.CustomerId);
      var address = getRowById_('CUSTOMER_ADDRESSES', 'AddressId', sub.AddressId);
      if (!customer || !address) return;
      var result = submitOrder({
        name: customer.Name, mobile: customer.Mobile, email: customer.Email, addressLine: address.AddressLine,
        area: address.Area, pin: address.PIN, lat: address.Lat, lng: address.Lng, slotId: sub.SlotId,
        cart: [{ productId: sub.ProductId, qty: sub.Qty }]
      });
      if (result.success) {
        updateRowFields_('SUBSCRIPTIONS', 'SubscriptionId', sub.SubscriptionId, { NextOrderDate: nextOrderDateFor_(sub.Frequency), LastOrderId: result.data.orderId });
        placed++;
      }
    } catch (e) {
      Logger.log('processDueSubscriptions failed for ' + sub.SubscriptionId + ': ' + e);
    }
  });
  return { placed: placed };
}

/* ============================================================================
 * SECTION 31G — RETURNS & REFUNDS
 * ========================================================================== */

function submitReturnRequest(orderId, reason, refundMethod) {
  return safeInvoke_('submitReturnRequest', function () {
    var order = getRowById_('ORDERS', 'OrderId', orderId);
    if (!order) return errorResponse_('Order not found.', 'NOT_FOUND');
    if (['DELIVERED', 'COMPLETED'].indexOf(order.Status) === -1) return errorResponse_('Returns can only be requested for delivered orders.', 'VALIDATION_ERROR');
    var returnId = generateId_('RET');
    appendRow_('RETURNS', {
      ReturnId: returnId, OrderId: orderId, CustomerId: order.CustomerId, Reason: safeString_(reason),
      RefundMethod: refundMethod === 'ORIGINAL' ? 'ORIGINAL' : 'WALLET', RefundAmount: 0, Status: 'REQUESTED',
      RequestedAt: nowIso_(), ResolvedAt: ''
    });
    writeAudit_(currentEmail_(), currentRole_(), 'RETURN_REQUESTED', 'RETURNS', returnId, '', reason, '');
    notify_('RETURN_REQUESTED_ADMIN', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), { orderId: orderId, reason: reason });
    return successResponse_({ returnId: returnId }, 'Return request submitted. Our team will review it shortly.');
  });
}

function getAdminReturns() {
  return safeInvoke_('getAdminReturns', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var returns = sheetToObjects_('RETURNS');
    returns.sort(function (a, b) { return new Date(b.RequestedAt) - new Date(a.RequestedAt); });
    return successResponse_(returns);
  });
}

function adminProcessReturn(returnId, approve, refundAmount) {
  return safeInvoke_('adminProcessReturn', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var ret = getRowById_('RETURNS', 'ReturnId', returnId);
    if (!ret) return errorResponse_('Return request not found.', 'NOT_FOUND');
    if (approve) {
      var amount = safeNumber_(refundAmount);
      if (amount > 0) {
        creditWallet_(ret.CustomerId, amount, 'Refund for order ' + ret.OrderId, ret.OrderId);
        updateRowFields_('ORDERS', 'OrderId', ret.OrderId, { PaymentStatus: 'REFUNDED' });
      }
      updateRowFields_('RETURNS', 'ReturnId', returnId, { Status: 'REFUNDED', RefundAmount: amount, ResolvedAt: nowIso_() });
      writeAudit_(currentEmail_(), currentRole_(), 'RETURN_APPROVED', 'RETURNS', returnId, '', amount, 'Refunded to wallet');
    } else {
      updateRowFields_('RETURNS', 'ReturnId', returnId, { Status: 'REJECTED', ResolvedAt: nowIso_() });
      writeAudit_(currentEmail_(), currentRole_(), 'RETURN_REJECTED', 'RETURNS', returnId, '', '', '');
    }
    return successResponse_({}, approve ? 'Return approved and refunded to wallet.' : 'Return request rejected.');
  });
}

/* ============================================================================
 * SECTION 31H — SUPPORT TICKETING (WITH ENHANCED CHAT THREADS)
 * ========================================================================== */

function submitSupportTicket(payload) {
  return safeInvoke_('submitSupportTicket', function () {
    if (!payload || !safeString_(payload.mobile) || !safeString_(payload.subject)) return errorResponse_('Mobile and subject are required.', 'VALIDATION_ERROR');
    var customerId = upsertCustomer_(payload.name, payload.mobile, payload.email);
    var ticketId = generateId_('TICKET');
    appendRow_('SUPPORT_TICKETS', {
      TicketId: ticketId, CustomerId: customerId, OrderId: safeString_(payload.orderId), Subject: safeString_(payload.subject),
      Category: TICKET_CATEGORIES.indexOf(payload.category) !== -1 ? payload.category : 'General Inquiry',
      Priority: payload.priority || 'NORMAL', Status: 'OPEN', CreatedAt: nowIso_(), UpdatedAt: nowIso_()
    });
    if (payload.message) {
      appendRow_('TICKET_MESSAGES', { MessageId: generateId_('TMSG'), TicketId: ticketId, Sender: payload.name || 'Customer', SenderRole: 'CUSTOMER', Message: safeString_(payload.message), Timestamp: nowIso_() });
    }
    writeAudit_(currentEmail_(), currentRole_(), 'TICKET_CREATED', 'SUPPORT_TICKETS', ticketId, '', payload.subject, '');
    notify_('SUPPORT_TICKET_CREATED_ADMIN', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), { ticketId: ticketId, subject: payload.subject });
    return successResponse_({ ticketId: ticketId }, 'Support ticket created. We will get back to you shortly.');
  });
}

function getMyTickets(mobile) {
  return safeInvoke_('getMyTickets', function () {
    if (!mobile) return successResponse_([]);
    var cleanSearchMobile = String(mobile).replace(/\D/g, '').trim();
    
    var customers = sheetToObjects_('CUSTOMERS');
    var matchingCustomerIds = [];
    customers.forEach(function (c) {
      var cMobile = String(c.Mobile || '').replace(/\D/g, '').trim();
      if (cMobile === cleanSearchMobile || cMobile.endsWith(cleanSearchMobile) || cleanSearchMobile.endsWith(cMobile)) {
        matchingCustomerIds.push(c.CustomerId);
      }
    });

    if (matchingCustomerIds.length === 0) return successResponse_([]);

    var tickets = sheetToObjects_('SUPPORT_TICKETS').filter(function (t) { return matchingCustomerIds.indexOf(t.CustomerId) !== -1; });
    var messages = sheetToObjects_('TICKET_MESSAGES');
    
    tickets.forEach(function (t) { 
      t.messages = messages.filter(function (m) { return m.TicketId === t.TicketId; });
      t.messages.sort(function (a, b) { return new Date(a.Timestamp) - new Date(b.Timestamp); });
    });
    
    tickets.sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return successResponse_(tickets);
  });
}

function addTicketMessage(ticketId, sender, message, isAdmin) {
  return safeInvoke_('addTicketMessage', function () {
    if (!safeString_(message)) return errorResponse_('Message cannot be empty.', 'VALIDATION_ERROR');
    var role = isAdmin ? 'STAFF' : 'CUSTOMER';
    var senderName = sender || (isAdmin ? 'DesiMurga Support' : 'Customer');
    
    appendRow_('TICKET_MESSAGES', { MessageId: generateId_('TMSG'), TicketId: ticketId, Sender: senderName, SenderRole: role, Message: safeString_(message), Timestamp: nowIso_() });
    updateRowFields_('SUPPORT_TICKETS', 'TicketId', ticketId, { UpdatedAt: nowIso_(), Status: isAdmin ? 'IN_PROGRESS' : 'OPEN' });
    return successResponse_({}, 'Message sent.');
  });
}

function getAdminTickets(statusFilter) {
  return safeInvoke_('getAdminTickets', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var tickets = sheetToObjects_('SUPPORT_TICKETS');
    var messages = sheetToObjects_('TICKET_MESSAGES');
    var customers = sheetToObjects_('CUSTOMERS');
    var customerMap = {};
    customers.forEach(function (c) { customerMap[c.CustomerId] = c; });
    
    tickets.forEach(function (t) { 
      t.messages = messages.filter(function (m) { return m.TicketId === t.TicketId; }); 
      t.messages.sort(function (a, b) { return new Date(a.Timestamp) - new Date(b.Timestamp); });
      t.customer = customerMap[t.CustomerId] || {}; 
    });
    
    if (statusFilter) tickets = tickets.filter(function (t) { return t.Status === statusFilter; });
    tickets.sort(function (a, b) { return new Date(b.UpdatedAt) - new Date(a.UpdatedAt); });
    return successResponse_(tickets);
  });
}

function adminReplyTicket(ticketId, message, closeTicket) {
  return safeInvoke_('adminReplyTicket', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    addTicketMessage(ticketId, 'DesiMurga Support', message, true);
    if (closeTicket) updateRowFields_('SUPPORT_TICKETS', 'TicketId', ticketId, { Status: 'RESOLVED', UpdatedAt: nowIso_() });
    writeAudit_(currentEmail_(), currentRole_(), 'TICKET_REPLIED', 'SUPPORT_TICKETS', ticketId, '', message, '');
    return successResponse_({}, 'Reply sent successfully.');
  });
}

function checkSupportTicketSla() {
  var slaHours = safeNumber_(getConfigValue_('supportTicketSlaHours', DEFAULT_CONFIG.supportTicketSlaHours));
  var openTickets = sheetToObjects_('SUPPORT_TICKETS').filter(function (t) { return t.Status === 'OPEN'; });
  var breached = openTickets.filter(function (t) { return (new Date() - new Date(t.CreatedAt)) / 3600000 > slaHours; });
  if (breached.length > 0) {
    notify_('SUPPORT_TICKET_SLA_BREACH', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), { count: breached.length });
  }
  return { breached: breached.length };
}

/* ============================================================================
 * SECTION 31I — MARKETING CAMPAIGNS & CUSTOMER SEGMENTATION
 * ========================================================================== */

function computeCustomerSegments_() {
  var customers = sheetToObjects_('CUSTOMERS').filter(function (c) { 
    return safeBoolean_(c.Active, true); 
  });
  
  var orders = sheetToObjects_('ORDERS').filter(function (o) { 
    return o.PaymentStatus === 'PAID' || o.Status === 'DELIVERED' || o.Status === 'COMPLETED'; 
  });
  
  var byCustomer = {};
  orders.forEach(function (o) {
    if (!byCustomer[o.CustomerId]) byCustomer[o.CustomerId] = { count: 0, total: 0, lastOrder: null };
    byCustomer[o.CustomerId].count++;
    byCustomer[o.CustomerId].total += safeNumber_(o.Total);
    var d = new Date(o.CreatedAt);
    if (!byCustomer[o.CustomerId].lastOrder || d > byCustomer[o.CustomerId].lastOrder) byCustomer[o.CustomerId].lastOrder = d;
  });

  var wholesaleIds = {};
  sheetToObjects_('WHOLESALE_ACCOUNTS').forEach(function (w) { 
    if (safeBoolean_(w.Active, true)) wholesaleIds[w.CustomerId] = true; 
  });

  var segments = { 
    ALL_CUSTOMERS: customers, 
    REPEAT_CUSTOMERS: [], 
    CHURN_RISK: [], 
    HIGH_VALUE: [], 
    WHOLESALE: [] 
  };

  customers.forEach(function (c) {
    var stats = byCustomer[c.CustomerId] || { count: 0, total: 0, lastOrder: null };
    if (stats.count >= 2) segments.REPEAT_CUSTOMERS.push(c);
    if (stats.total >= 1500) segments.HIGH_VALUE.push(c);
    if (wholesaleIds[c.CustomerId]) segments.WHOLESALE.push(c);
    if (stats.lastOrder && (new Date() - stats.lastOrder) / 86400000 > 14 && stats.count > 0) segments.CHURN_RISK.push(c);
  });

  return segments;
}

function getAdminSegmentRecipients(segmentKey, adminToken) {
  return safeInvoke_('getAdminSegmentRecipients', function () {
    requireRole_(['ADMIN', 'MANAGER'], adminToken);
    var segments = computeCustomerSegments_();
    var list = segments[segmentKey] || segments.ALL_CUSTOMERS;
    return successResponse_(list);
  });
}

function adminCreateCampaign(payload, adminToken) {
  return safeInvoke_('adminCreateCampaign', function () {
    requireRole_(['ADMIN', 'MANAGER'], adminToken);
    if (!payload || !safeString_(payload.name) || !safeString_(payload.messageTemplate)) {
      return errorResponse_('Campaign title and message copy are required.', 'VALIDATION_ERROR');
    }
    
    var campaignId = generateId_('CAMPAIGN');
    appendRow_('CAMPAIGNS', {
      CampaignId: campaignId,
      Name: safeString_(payload.name),
      Segment: payload.segment || 'ALL_CUSTOMERS',
      Channel: safeString_(payload.channel) || 'WHATSAPP_EMAIL',
      MessageTemplate: safeString_(payload.messageTemplate),
      Status: 'DRAFT',
      ScheduledAt: '',
      SentCount: 0,
      CreatedAt: nowIso_()
    });
    
    writeAudit_(currentEmail_(), 'ADMIN', 'CAMPAIGN_CREATED', 'CAMPAIGNS', campaignId, '', payload.segment, 'Created campaign ' + payload.name);
    return successResponse_({ campaignId: campaignId }, 'Campaign created successfully.');
  });
}

function adminSendCampaign(campaignId, adminToken) {
  return safeInvoke_('adminSendCampaign', function () {
    requireRole_(['ADMIN', 'MANAGER'], adminToken);
    var campaign = getRowById_('CAMPAIGNS', 'CampaignId', campaignId);
    if (!campaign) return errorResponse_('Campaign record not found.', 'NOT_FOUND');
    
    var segments = computeCustomerSegments_();
    var recipients = segments[campaign.Segment] || [];
    
    // Auto-fallback to ALL_CUSTOMERS if segment is empty
    if (recipients.length === 0) {
      recipients = segments.ALL_CUSTOMERS;
    }
    
    var validRecipients = recipients.filter(function (c) {
      return c.Email && String(c.Email).indexOf('@') > -1 && String(c.Email).indexOf('.') > -1;
    });

    var brandName = getConfigValue_('brandName', DEFAULT_CONFIG.brandName);
    var subjectLine = brandName + ' — ' + campaign.Name;
    var sentCount = 0;
    var skippedCount = recipients.length - validRecipients.length;

    if (validRecipients.length === 0) {
      // Send preview copy to Admin
      var adminEmail = getConfigValue_('supportEmail', Session.getEffectiveUser().getEmail());
      if (adminEmail && adminEmail.indexOf('@') > -1) {
        var dummyCustomer = { Name: 'Customer Preview', Email: adminEmail };
        GmailApp.sendEmail(adminEmail, '[PREVIEW] ' + subjectLine, '', {
          htmlBody: buildCampaignBody_(campaign, dummyCustomer),
          name: brandName
        });
      }
      return errorResponse_('None of the customers in this segment have an email on file. Use the 1-Click WhatsApp buttons instead!', 'NO_EMAILS');
    }

    validRecipients.forEach(function (c) {
      try {
        GmailApp.sendEmail(String(c.Email).trim(), subjectLine, '', {
          htmlBody: buildCampaignBody_(campaign, c),
          name: brandName
        });
        sentCount++;
      } catch (e) {
        Logger.log('Campaign send failed for ' + c.Email + ': ' + e);
      }
    });

    updateRowFields_('CAMPAIGNS', 'CampaignId', campaignId, {
      Status: 'SENT',
      ScheduledAt: nowIso_(),
      SentCount: sentCount
    });

    return successResponse_({ 
      sentCount: sentCount, 
      skippedCount: skippedCount 
    }, 'Campaign emailed to ' + sentCount + ' customers (' + skippedCount + ' without email). Check your Gmail Sent box!');
  });
}

function adminDeleteCampaign(campaignId) {
  return safeInvoke_('adminDeleteCampaign', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    deleteRow_('CAMPAIGNS', 'CampaignId', campaignId);
    return successResponse_({}, 'Campaign deleted.');
  });
}

function buildCampaignBody_(campaign, customer) {
  var brandName = esc_(getConfigValue_('brandName', DEFAULT_CONFIG.brandName));
  var tagline = esc_(getConfigValue_('tagline', DEFAULT_CONFIG.tagline));
  var locationName = esc_(getConfigValue_('locationName', DEFAULT_CONFIG.locationName));
  var supportPhone = esc_(getConfigValue_('supportPhone', DEFAULT_CONFIG.supportPhone));
  var logoUrl = getConfigValue_('logoUrl', '');
  var directLogoUrl = logoUrl ? ('https://lh3.googleusercontent.com/d/' + extractDriveId_(logoUrl)) : '';
  var webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) {}

  var personalizedText = safeString_(campaign.MessageTemplate).replace(/\{name\}/g, esc_(customer.Name || 'Food Lover'));

  return '<!DOCTYPE html>' +
  '<html>' +
  '<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>' +
  '<body style="margin:0;padding:0;background-color:#f5ece0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">' +
    '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#f5ece0;padding:24px 12px;">' +
      '<tr><td align="center">' +
        '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width:540px;background-color:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 6px 24px rgba(90,60,30,0.08);border:1px solid #ebdccb;">' +
          
          // Header Banner
          '<tr><td style="background:linear-gradient(135deg, #7a4a2b 0%, #5c3620 100%);padding:26px 24px;text-align:center;">' +
            (directLogoUrl ? '<img src="' + esc_(directLogoUrl) + '" alt="DesiMurga" width="56" height="56" style="border-radius:50%;border:2px solid rgba(255,255,255,0.3);margin-bottom:8px;object-fit:cover;">' : '<div style="font-size:32px;margin-bottom:6px;">🐓</div>') +
            '<h1 style="color:#ffffff;font-size:22px;font-weight:900;margin:0;">' + brandName + '</h1>' +
            '<p style="color:#f5ece0;font-size:11px;margin:3px 0 0;opacity:0.85;">' + tagline + '</p>' +
          '</td></tr>' +

          // Campaign Content
          '<tr><td style="padding:28px 24px;">' +
            '<div style="display:inline-block;padding:4px 12px;background:#fbf0dc;color:#c98a2c;border-radius:20px;font-size:11px;font-weight:800;margin-bottom:12px;">EXCLUSIVE OFFER</div>' +
            '<h2 style="color:#5c3620;font-size:20px;font-weight:800;margin:0 0 14px;line-height:1.3;">' + esc_(campaign.Name) + '</h2>' +
            '<div style="font-size:14px;color:#5c4a3a;line-height:1.7;white-space:pre-line;margin-bottom:24px;background:#fdf8f2;padding:18px;border-radius:14px;border:1px solid #f0e6d8;">' +
              personalizedText +
            '</div>' +
            '<div style="text-align:center;margin:10px 0;">' +
              '<a href="' + (webAppUrl || '#') + '" target="_blank" style="display:inline-block;background:#7a4a2b;color:#ffffff;text-decoration:none;font-size:14px;font-weight:800;padding:14px 32px;border-radius:12px;box-shadow:0 3px 12px rgba(122,74,43,0.3);">Order Fresh Desi Chicken &rarr;</a>' +
            '</div>' +
          '</td></tr>' +

          // Footer
          '<tr><td style="background:#2b211a;color:#d9c7b3;padding:20px;text-align:center;font-size:11px;line-height:1.6;">' +
            '<p style="margin:0 0 4px;color:#ffffff;font-weight:700;">DesiMurga™ Kolkata</p>' +
            '<p style="margin:0;color:#a89888;">Delivering fresh across ' + locationName + ' · Call/WhatsApp: ' + supportPhone + '</p>' +
          '</td></tr>' +

        '</table>' +
      '</td></tr>' +
    '</table>' +
  '</body></html>';
}

function getAdminCampaigns() {
  return safeInvoke_('getAdminCampaigns', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(sheetToObjects_('CAMPAIGNS'));
  });
}

function getAdminSegmentCounts() {
  return safeInvoke_('getAdminSegmentCounts', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var segments = computeCustomerSegments_();
    var counts = {};
    Object.keys(segments).forEach(function (k) { counts[k] = segments[k].length; });
    return successResponse_(counts);
  });
}

/* ============================================================================
 * SECTION 31J — VENDOR PAYMENTS (ACCOUNTS PAYABLE)
 * ========================================================================== */

function adminRecordVendorPayment(payload) {
  return safeInvoke_('adminRecordVendorPayment', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    if (!payload || !payload.supplierId || !safeNumber_(payload.amount)) return errorResponse_('Supplier and amount are required.', 'VALIDATION_ERROR');
    var paymentId = generateId_('VPAY');
    appendRow_('VENDOR_PAYMENTS', {
      PaymentId: paymentId, SupplierId: payload.supplierId, PoId: safeString_(payload.poId), Amount: safeNumber_(payload.amount),
      Method: safeString_(payload.method) || 'Bank Transfer', Status: 'PAID', PaidAt: nowIso_(), Notes: safeString_(payload.notes)
    });
    writeAudit_(currentEmail_(), currentRole_(), 'VENDOR_PAYMENT_RECORDED', 'VENDOR_PAYMENTS', paymentId, '', payload.amount, 'Payment to supplier ' + payload.supplierId);
    return successResponse_({ paymentId: paymentId }, 'Vendor payment recorded.');
  });
}

function getAdminVendorPayments() {
  return safeInvoke_('getAdminVendorPayments', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    var payments = sheetToObjects_('VENDOR_PAYMENTS');
    var suppliers = sheetToObjects_('SUPPLIERS');
    var supplierMap = {};
    suppliers.forEach(function (s) { supplierMap[s.SupplierId] = s.BusinessName; });
    payments.forEach(function (p) { p.supplierName = supplierMap[p.SupplierId] || 'Unknown'; });
    payments.sort(function (a, b) { return new Date(b.PaidAt) - new Date(a.PaidAt); });
    return successResponse_(payments);
  });
}

function getAccountsPayableSummary_() {
  var pos = sheetToObjects_('PURCHASE_ORDERS').filter(function (p) { return p.Status === 'RECEIVED'; });
  var payments = sheetToObjects_('VENDOR_PAYMENTS');
  var paidByPo = {};
  payments.forEach(function (p) { paidByPo[p.PoId] = (paidByPo[p.PoId] || 0) + safeNumber_(p.Amount); });
  var totalOwed = 0, totalPaid = 0;
  pos.forEach(function (po) {
    totalOwed += safeNumber_(po.EstimatedTotal);
    totalPaid += (paidByPo[po.PoId] || 0);
  });
  return { totalOwed: Math.round(totalOwed * 100) / 100, totalPaid: Math.round(totalPaid * 100) / 100, outstanding: Math.round((totalOwed - totalPaid) * 100) / 100 };
}

function getAdminAccountsPayable() {
  return safeInvoke_('getAdminAccountsPayable', function () {
    requireRole_(['ADMIN', 'MANAGER', 'PROCUREMENT']);
    return successResponse_(getAccountsPayableSummary_());
  });
}

/* ============================================================================
 * SECTION 31K — ADVANCED CRM ANALYTICS (RFM, LTV, COHORTS)
 * ========================================================================== */

function getAdminCustomerAnalytics() {
  return safeInvoke_('getAdminCustomerAnalytics', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var customers = sheetToObjects_('CUSTOMERS');
    var orders = sheetToObjects_('ORDERS').filter(function (o) { return o.PaymentStatus === 'PAID'; });
    var byCustomer = {};
    orders.forEach(function (o) {
      if (!byCustomer[o.CustomerId]) byCustomer[o.CustomerId] = { orders: 0, revenue: 0, lastOrder: null, firstOrder: null };
      var stat = byCustomer[o.CustomerId];
      stat.orders++;
      stat.revenue += safeNumber_(o.Total);
      var d = new Date(o.CreatedAt);
      if (!stat.lastOrder || d > stat.lastOrder) stat.lastOrder = d;
      if (!stat.firstOrder || d < stat.firstOrder) stat.firstOrder = d;
    });
    var rows = customers.map(function (c) {
      var stat = byCustomer[c.CustomerId] || { orders: 0, revenue: 0, lastOrder: null, firstOrder: null };
      var recencyDays = stat.lastOrder ? Math.floor((new Date() - stat.lastOrder) / 86400000) : null;
      var ltv = Math.round(stat.revenue * 100) / 100;
      var rfmScore = (recencyDays !== null && recencyDays < 14 ? 3 : recencyDays !== null && recencyDays < 30 ? 2 : 1) +
        (stat.orders >= 5 ? 3 : stat.orders >= 2 ? 2 : stat.orders >= 1 ? 1 : 0) +
        (ltv >= 3000 ? 3 : ltv >= 1000 ? 2 : ltv > 0 ? 1 : 0);
      return { customerId: c.CustomerId, name: c.Name, mobile: c.Mobile, orders: stat.orders, revenue: ltv, recencyDays: recencyDays, rfmScore: rfmScore };
    });
    rows.sort(function (a, b) { return b.revenue - a.revenue; });
    var avgLtv = rows.length ? Math.round((rows.reduce(function (s, r) { return s + r.revenue; }, 0) / rows.length) * 100) / 100 : 0;
    var repeatRate = rows.length ? Math.round((rows.filter(function (r) { return r.orders >= 2; }).length / rows.length) * 10000) / 100 : 0;
    return successResponse_({ customers: rows.slice(0, 100), avgLtv: avgLtv, repeatRate: repeatRate, totalCustomers: rows.length });
  });
}

/* ============================================================================
 * SECTION 31L — BATCH TRACEABILITY (QR CODES)
 * ========================================================================== */

function getBatchTraceability(batchId) {
  return safeInvoke_('getBatchTraceability', function () {
    var batch = getRowById_('LIVE_BATCHES', 'BatchId', batchId);
    if (!batch) return errorResponse_('Batch not found.', 'NOT_FOUND');
    var receiving = getRowById_('RECEIVING', 'ReceivingId', batch.ReceivingId);
    var supplier = getRowById_('SUPPLIERS', 'SupplierId', batch.SupplierId);
    var qc = sheetToObjects_('QC_RECORDS').filter(function (q) { return q.BatchId === batchId; });
    var processing = sheetToObjects_('PROCESSING_BATCHES').filter(function (p) { return p.SourceLiveBatchId === batchId; });
    var po = receiving ? getRowById_('PURCHASE_ORDERS', 'PoId', receiving.PoId) : null;
    return successResponse_({ batch: batch, receiving: receiving, supplier: supplier, qcRecords: qc, processingBatches: processing, purchaseOrder: po });
  });
}

function getBatchQrUrl_(batchId) {
  var baseUrl = ScriptApp.getService().getUrl();
  var traceUrl = baseUrl ? (baseUrl + '?trace=' + encodeURIComponent(batchId)) : batchId;
  return 'https://chart.googleapis.com/chart?cht=qr&chs=220x220&chl=' + encodeURIComponent(traceUrl);
}

function getAdminBatchQr(batchId) {
  return safeInvoke_('getAdminBatchQr', function () {
    requireRole_(['ADMIN', 'MANAGER', 'FARM', 'PROCESSING']);
    return successResponse_({ qrUrl: getBatchQrUrl_(batchId) });
  });
}

/* ============================================================================
 * SECTION 31M — PROFIT & LOSS STATEMENT
 * ========================================================================== */

function getProfitLossStatement(fromDate, toDate) {
  return safeInvoke_('getProfitLossStatement', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var orders = sheetToObjects_('ORDERS').filter(function (o) { return o.PaymentStatus === 'PAID'; });
    if (fromDate) orders = orders.filter(function (o) { return new Date(o.CreatedAt) >= new Date(fromDate); });
    if (toDate) orders = orders.filter(function (o) { return new Date(o.CreatedAt) <= new Date(toDate); });
    var revenue = orders.reduce(function (s, o) { return s + safeNumber_(o.Total); }, 0);
    var discounts = orders.reduce(function (s, o) { return s + safeNumber_(o.Discount); }, 0);
    var pos = sheetToObjects_('PURCHASE_ORDERS').filter(function (p) { return p.Status === 'RECEIVED'; });
    var cogs = pos.reduce(function (s, p) { return s + safeNumber_(p.EstimatedTotal); }, 0);
    var expenses = sheetToObjects_('EXPENSES');
    var expenseByCategory = {};
    var totalExpenses = 0;
    expenses.forEach(function (e) { expenseByCategory[e.Category] = (expenseByCategory[e.Category] || 0) + safeNumber_(e.Amount); totalExpenses += safeNumber_(e.Amount); });
    var walletCredits = sheetToObjects_('WALLET_TRANSACTIONS').filter(function (t) { return t.Type === 'CREDIT' && ['Referral', 'Redeemed', 'Refund'].some(function (k) { return t.Reason.indexOf(k) !== -1; }); });
    var loyaltyReferralCost = walletCredits.reduce(function (s, t) { return s + safeNumber_(t.Amount); }, 0);
    var grossProfit = Math.round((revenue - cogs) * 100) / 100;
    var netProfit = Math.round((grossProfit - totalExpenses - loyaltyReferralCost) * 100) / 100;
    return successResponse_({
      revenue: Math.round(revenue * 100) / 100, discounts: Math.round(discounts * 100) / 100, cogs: Math.round(cogs * 100) / 100,
      grossProfit: grossProfit, expenseByCategory: expenseByCategory, totalExpenses: Math.round(totalExpenses * 100) / 100,
      loyaltyReferralCost: Math.round(loyaltyReferralCost * 100) / 100, netProfit: netProfit,
      netMargin: revenue > 0 ? Math.round((netProfit / revenue) * 10000) / 100 : 0
    });
  });
}

/* ============================================================================
 * SECTION 31N — ABANDONED CART RECOVERY
 * ========================================================================== */

function logAbandonedCart(mobile, email, cartLines) {
  return safeInvoke_('logAbandonedCart', function () {
    if (!email || email.indexOf('@') === -1) return successResponse_({}, 'No email supplied; skipping.');
    var summary = (cartLines || []).map(function (l) { return l.name + ' x' + l.qty; }).join(', ');
    notify_('ABANDONED_CART', email, { summary: summary });
    return successResponse_({}, 'Reminder scheduled.');
  }, { silent: true });
}

/* ============================================================================
 * SECTION 31O — WISHLIST / FAVOURITES
 * ========================================================================== */

function toggleWishlist(mobile, name, productId) {
  return safeInvoke_('toggleWishlist', function () {
    if (!mobile || !productId) return errorResponse_('Mobile number and product are required.', 'VALIDATION_ERROR');
    var customerId = upsertCustomer_(name || 'Guest', mobile, '');
    var existing = sheetToObjects_('WISHLIST').filter(function (w) { return w.CustomerId === customerId && w.ProductId === productId; })[0];
    if (existing) {
      deleteRow_('WISHLIST', 'WishlistId', existing.WishlistId);
      return successResponse_({ saved: false }, 'Removed from wishlist.');
    }
    appendRow_('WISHLIST', { WishlistId: generateId_('WISH'), CustomerId: customerId, ProductId: productId, CreatedAt: nowIso_() });
    return successResponse_({ saved: true }, 'Saved to wishlist.');
  });
}

function getMyWishlist(mobile) {
  return safeInvoke_('getMyWishlist', function () {
    if (!mobile) return successResponse_({ productIds: [], products: [] });
    var customer = sheetToObjects_('CUSTOMERS').filter(function (c) { return c.Mobile === mobile; })[0];
    if (!customer) return successResponse_({ productIds: [], products: [] });
    var wish = sheetToObjects_('WISHLIST').filter(function (w) { return w.CustomerId === customer.CustomerId; });
    var productIds = wish.map(function (w) { return w.ProductId; });
    var products = sheetToObjects_('PRODUCTS').filter(function (p) { return productIds.indexOf(p.ProductId) !== -1; });
    return successResponse_({ productIds: productIds, products: products });
  });
}

/* ============================================================================
 * SECTION 31P — STATIC CONTENT PAGES
 * ========================================================================== */

function getStaticPage(pageKey) {
  return safeInvoke_('getStaticPage', function () {
    var page = getRowById_('STATIC_PAGES', 'PageKey', pageKey);
    if (!page) return errorResponse_('Page not found.', 'NOT_FOUND');
    return successResponse_(page);
  });
}

function getAllStaticPagesAdmin() {
  return safeInvoke_('getAllStaticPagesAdmin', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    return successResponse_(sheetToObjects_('STATIC_PAGES'));
  });
}

function adminSaveStaticPage(pageKey, title, content) {
  return safeInvoke_('adminSaveStaticPage', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var existing = getRowById_('STATIC_PAGES', 'PageKey', pageKey);
    var fields = { Title: safeString_(title), Content: safeString_(content), UpdatedAt: nowIso_() };
    if (existing) updateRowFields_('STATIC_PAGES', 'PageKey', pageKey, fields);
    else appendRow_('STATIC_PAGES', Object.assign({ PageKey: pageKey }, fields));
    writeAudit_(currentEmail_(), currentRole_(), 'STATIC_PAGE_UPDATED', 'STATIC_PAGES', pageKey, '', '', 'Content page updated');
    return successResponse_({}, 'Page updated.');
  });
}

/* ============================================================================
 * SECTION 31Q — FULL-SYSTEM QA TEST HARNESS
 * ========================================================================== */

var QA_TEST_MOBILE = '9000000000';
var QA_MARKER = 'QA-TEST';

function runFullSystemTest_() {
  var report = [];
  var created = { PRODUCTS: [], SUPPLIERS: [], PURCHASE_ORDERS: [], RECEIVING: [], LIVE_BATCHES: [], QC_RECORDS: [],
    ORDERS: [], ORDER_LINES: [], PAYMENTS: [], INVOICES: [], PROCESSING_BATCHES: [], DRESSED_INVENTORY: [],
    CUSTOMERS: [], CUSTOMER_ADDRESSES: [], COUPONS: [], GIFT_CARDS: [], SUBSCRIPTIONS: [], RETURNS: [],
    SUPPORT_TICKETS: [], TICKET_MESSAGES: [], WALLET_TRANSACTIONS: [], LOYALTY: [], REFERRALS: [], DELIVERIES: [] };

  function step(name, fn) {
    try {
      var result = fn();
      report.push({ step: name, status: 'PASS', detail: result || 'OK' });
      return result;
    } catch (e) {
      report.push({ step: name, status: 'FAIL', detail: String(e && e.message ? e.message : e) });
      return null;
    }
  }

  step('Database health check', function () {
    var health = runDatabaseHealthCheck();
    if (!health.healthy) throw new Error('Unhealthy: ' + health.issues.join('; '));
    return 'All ' + Object.keys(SCHEMA).length + ' sheets present and valid.';
  });

  var testProductId = step('Create test product', function () {
    var id = generateId_('PRD');
    appendRow_('PRODUCTS', {
      ProductId: id, Name: QA_MARKER + ' Product', Description: 'QA harness test product', Category: 'Whole Desi Hen',
      ImageUrl: '', Price: 200, Unit: 'per bird', EstWeight: 1, MinQty: 1, MaxQty: 10, Available: true, Featured: false,
      Active: true, BreedType: 'Desi', BirdType: 'Hen', Source: 'QA', AgeRange: '', TypicalWeight: '', FarmingMethod: '', CreatedAt: nowIso_()
    });
    created.PRODUCTS.push(id);
    return id;
  });

  var testSupplierId = step('Create test supplier', function () {
    var id = generateId_('SUP');
    appendRow_('SUPPLIERS', { SupplierId: id, BusinessName: QA_MARKER + ' Supplier', ContactPerson: 'QA', Phone: '', WhatsApp: '', Address: '', Location: '', BreedType: 'Desi', MinOrder: 1, Price: 250, AvgWeight: 1.3, Capacity: 100, LeadTime: 1, Reliability: 'Good', Active: true, Notes: QA_MARKER, CreatedAt: nowIso_() });
    created.SUPPLIERS.push(id);
    return id;
  });
  var testPoId = step('Create test purchase order', function () {
    if (!testSupplierId) throw new Error('Skipped — supplier step failed.');
    var id = generateId_('PO');
    appendRow_('PURCHASE_ORDERS', { PoId: id, SupplierId: testSupplierId, Date: nowIso_(), RequiredDate: nowIso_(), Product: 'Desi Hen (Live)', Quantity: 5, TargetWeight: 6.5, AgreedPrice: 250, EstimatedTotal: 1250, Status: 'SENT', Notes: QA_MARKER, CreatedAt: nowIso_() });
    created.PURCHASE_ORDERS.push(id);
    return id;
  });
  var testBatchId = step('Receive test purchase order (creates live batch)', function () {
    if (!testPoId) throw new Error('Skipped — PO step failed.');
    var receivingId = generateId_('RECV');
    appendRow_('RECEIVING', { ReceivingId: receivingId, PoId: testPoId, QtyOrdered: 5, QtyReceived: 5, Rejected: 0, Mortality: 0, Accepted: 5, AvgLiveWeight: 1.3, TotalLiveWeight: 6.5, ReceivingDate: nowIso_(), QcStatus: 'PASSED' });
    created.RECEIVING.push(receivingId);
    var batchId = generateId_('BATCH');
    appendRow_('LIVE_BATCHES', { BatchId: batchId, SupplierId: testSupplierId, ReceivingId: receivingId, BirdType: 'Hen', Sex: 'Mixed', Breed: 'Desi', QtyReceived: 5, QtyAvailable: 5, QtyReserved: 0, QtyProcessing: 0, QtyProcessed: 0, Mortality: 0, AvgLiveWeight: 1.3, TotalLiveWeight: 6.5, ArrivalDate: nowIso_(), ExpectedProcessingDate: '', Status: 'ACTIVE' });
    created.LIVE_BATCHES.push(batchId);
    updateRowFields_('PURCHASE_ORDERS', 'PoId', testPoId, { Status: 'RECEIVED' });
    var qcId = generateId_('QC');
    appendRow_('QC_RECORDS', { QcId: qcId, BatchId: batchId, SupplierId: testSupplierId, ReceivedQty: 5, Healthy: 5, Rejected: 0, Mortality: 0, Condition: 'Good', WeightCheck: 'OK', Notes: QA_MARKER, Inspector: 'QA', Timestamp: nowIso_(), Status: 'PASSED' });
    created.QC_RECORDS.push(qcId);
    recalcLiveInventorySummary_();
    return batchId;
  });

  step('Create test coupon and validate it', function () {
    var couponId = generateId_('COUPON');
    appendRow_('COUPONS', { CouponId: couponId, Code: 'QATEST10', DiscountType: 'FLAT', Value: 10, MinOrder: 0, ValidFrom: '', ValidTo: '', Active: true, UsageLimit: 100, UsedCount: 0 });
    created.COUPONS.push(couponId);
    var result = evaluateCoupon_('QATEST10', 200);
    if (!result.valid) throw new Error('Coupon evaluated as invalid.');
    return 'Coupon validated, discount = ' + result.discount;
  });
  var giftCardCode = step('Issue and preview test gift card', function () {
    var cardId = generateId_('GIFT');
    var code = 'QA-GIFT-' + Utilities.getUuid().split('-')[0].toUpperCase();
    appendRow_('GIFT_CARDS', { CardId: cardId, Code: code, InitialValue: 50, Balance: 50, IssuedTo: QA_MARKER, ExpiresAt: '', Active: true, CreatedAt: nowIso_() });
    created.GIFT_CARDS.push(cardId);
    var preview = previewGiftCard_(code);
    if (!preview.valid || preview.balance !== 50) throw new Error('Gift card preview mismatch.');
    return code;
  });

  var testOrderId = step('Submit test order (cart + checkout engine)', function () {
    if (!testProductId) throw new Error('Skipped — product step failed.');
    var slots = sheetToObjects_('DELIVERY_SLOTS').filter(function (s) { return safeBoolean_(s.Active); });
    if (slots.length === 0) throw new Error('No active delivery slots configured.');
    var result = submitOrder({
      name: QA_MARKER + ' Customer', mobile: QA_TEST_MOBILE, email: 'qa-test@example.com',
      addressLine: 'QA Test Address', area: 'Garia', pin: '700084', slotId: slots[0].SlotId,
      cart: [{ productId: testProductId, qty: 1 }]
    });
    if (!result.success) throw new Error(result.message);
    return result.data.orderId;
  });
  step('Track created customer/order/lines for cleanup', function () {
    if (!testOrderId) return 'Skipped.';
    var order = getRowById_('ORDERS', 'OrderId', testOrderId);
    created.ORDERS.push(testOrderId);
    created.CUSTOMERS.push(order.CustomerId);
    var addr = getRowById_('CUSTOMER_ADDRESSES', 'AddressId', order.AddressId);
    if (addr) created.CUSTOMER_ADDRESSES.push(addr.AddressId);
    sheetToObjects_('ORDER_LINES').filter(function (l) { return l.OrderId === testOrderId; }).forEach(function (l) { created.ORDER_LINES.push(l.LineId); });
    return 'Tracked.';
  });
  step('Submit and verify test payment', function () {
    if (!testOrderId) throw new Error('Skipped — order step failed.');
    var payResult = submitPaymentReference(testOrderId, 'QA-UPI-REF-0001', getRowById_('ORDERS', 'OrderId', testOrderId).Total);
    if (!payResult.success) throw new Error(payResult.message);
    created.PAYMENTS.push(payResult.data.paymentId);
    var verifyResult = adminVerifyPayment(payResult.data.paymentId, true);
    if (!verifyResult.success) throw new Error(verifyResult.message);
    var order = getRowById_('ORDERS', 'OrderId', testOrderId);
    if (order.Status !== 'RESERVED') throw new Error('Expected RESERVED, got ' + order.Status);
    return 'Payment verified; live inventory reserved.';
  });
  step('Run a processing batch against the test order\u2019s reservation', function () {
    if (!testBatchId) throw new Error('Skipped — batch step failed.');
    var result = adminCreateProcessingBatch({ sourceLiveBatchId: testBatchId, birdCount: 1, totalLiveWeight: 1.3, rejected: 0, mortality: 0, dressedWeight: 0.9, saleableWeight: 0.9, productId: testProductId, cut: 'Whole' });
    if (!result.success) throw new Error(result.message);
    created.PROCESSING_BATCHES.push(result.data.procBatchId);
    return 'Yield ' + result.data.yieldPct + '%';
  });
  step('Advance order to Packed and generate invoice', function () {
    if (!testOrderId) throw new Error('Skipped — order step failed.');
    var r1 = adminTransitionOrder(testOrderId, 'PROCESSING');
    if (!r1.success) throw new Error(r1.message);
    var r2 = adminTransitionOrder(testOrderId, 'PACKED');
    if (!r2.success) throw new Error(r2.message);
    var invoice = sheetToObjects_('INVOICES').filter(function (i) { return i.OrderId === testOrderId; })[0];
    if (!invoice) throw new Error('Invoice was not generated.');
    created.INVOICES.push(invoice.InvoiceId);
    return 'Invoice ' + invoice.InvoiceId + ' generated at ' + invoice.PdfUrl;
  });
  step('Assign delivery and mark delivered', function () {
    if (!testOrderId) throw new Error('Skipped — order step failed.');
    var r1 = adminTransitionOrder(testOrderId, 'READY_FOR_DELIVERY');
    if (!r1.success) throw new Error(r1.message);
    var assignResult = adminAssignDelivery(testOrderId, 'QA Rider');
    if (!assignResult.success) throw new Error(assignResult.message);
    var delivery = sheetToObjects_('DELIVERIES').filter(function (d) { return d.OrderId === testOrderId; })[0];
    if (delivery) created.DELIVERIES.push(delivery.DeliveryId);
    var r2 = adminTransitionOrder(testOrderId, 'DELIVERED');
    if (!r2.success) throw new Error(r2.message);
    return 'Delivered.';
  });
  step('Complete order and verify loyalty points were earned', function () {
    if (!testOrderId) throw new Error('Skipped — order step failed.');
    var r = adminTransitionOrder(testOrderId, 'COMPLETED');
    if (!r.success) throw new Error(r.message);
    var order = getRowById_('ORDERS', 'OrderId', testOrderId);
    var loyalty = getRowById_('LOYALTY', 'CustomerId', order.CustomerId);
    if (!loyalty || safeNumber_(loyalty.Points) <= 0) throw new Error('No loyalty points were credited.');
    created.LOYALTY.push(loyalty.CustomerId);
    return safeNumber_(loyalty.Points) + ' points credited.';
  });

  step('Wallet credit/debit round-trip', function () {
    var order = getRowById_('ORDERS', 'OrderId', testOrderId);
    if (!order) throw new Error('Skipped — no test customer available.');
    var before = getWalletBalance_(order.CustomerId);
    creditWallet_(order.CustomerId, 25, QA_MARKER + ' wallet test credit', '');
    var after = getWalletBalance_(order.CustomerId);
    if (after !== before + 25) throw new Error('Wallet balance did not update correctly.');
    return 'Wallet balance moved from ' + before + ' to ' + after + '.';
  });
  step('Referral code generation', function () {
    var order = getRowById_('ORDERS', 'OrderId', testOrderId);
    if (!order) throw new Error('Skipped — no test customer available.');
    var code = ensureReferralCode_(order.CustomerId);
    if (!code) throw new Error('No referral code generated.');
    return 'Referral code: ' + code;
  });
  step('Create and cancel a test subscription', function () {
    if (!testProductId) throw new Error('Skipped — product step failed.');
    var slots = sheetToObjects_('DELIVERY_SLOTS').filter(function (s) { return safeBoolean_(s.Active); });
    var result = createSubscription({ name: QA_MARKER + ' Sub', mobile: QA_TEST_MOBILE, email: '', productId: testProductId, qty: 1, frequency: 'WEEKLY', addressLine: 'QA Address', area: 'Garia', pin: '700084', slotId: slots[0].SlotId });
    if (!result.success) throw new Error(result.message);
    created.SUBSCRIPTIONS.push(result.data.subscriptionId);
    updateSubscriptionStatus(result.data.subscriptionId, 'CANCELLED');
    return 'Subscription created and cancelled cleanly.';
  });
  step('Submit a test return request', function () {
    if (!testOrderId) throw new Error('Skipped — order step failed.');
    var result = submitReturnRequest(testOrderId, QA_MARKER + ' reason', 'WALLET');
    if (!result.success) throw new Error(result.message);
    created.RETURNS.push(result.data.returnId);
    return result.data.returnId;
  });
  step('Submit a test support ticket and reply', function () {
    var result = submitSupportTicket({ mobile: QA_TEST_MOBILE, name: QA_MARKER, email: '', orderId: testOrderId || '', category: 'General Inquiry', subject: QA_MARKER + ' ticket', message: 'Automated QA message' });
    if (!result.success) throw new Error(result.message);
    created.SUPPORT_TICKETS.push(result.data.ticketId);
    var msgs = sheetToObjects_('TICKET_MESSAGES').filter(function (m) { return m.TicketId === result.data.ticketId; });
    msgs.forEach(function (m) { created.TICKET_MESSAGES.push(m.MessageId); });
    return result.data.ticketId;
  });

  step('Customer segmentation calculation', function () {
    var segments = computeCustomerSegments_();
    return Object.keys(segments).map(function (k) { return k + ':' + segments[k].length; }).join(', ');
  });
  step('Procurement recommendation calculation', function () {
    var rec = computeProcurementRecommendation_();
    return 'Available ' + rec.available + ', recommended purchase ' + rec.recommendedPurchase;
  });
  step('Delivery radius validation logic', function () {
    var inside = isWithinDeliveryRadius_(getConfigValue_('deliveryCenterLat'), getConfigValue_('deliveryCenterLng'));
    if (!inside.withinRadius) throw new Error('Farm center coordinates evaluated as outside their own delivery radius.');
    var outside = isWithinDeliveryRadius_(0, 0);
    if (outside.withinRadius) throw new Error('Null Island incorrectly evaluated as inside delivery radius.');
    return 'Radius validation correct in both directions.';
  });
  step('P&L statement calculation', function () {
    var pl = getProfitLossStatement(null, null);
    if (!pl.success) throw new Error(pl.message);
    return 'Net profit: ' + pl.data.netProfit;
  });

  var passCount = report.filter(function (r) { return r.status === 'PASS'; }).length;
  return { steps: report, passCount: passCount, failCount: report.length - passCount, totalSteps: report.length, createdRecordIds: created };
}

function adminRunFullSystemTest() {
  return safeInvoke_('adminRunFullSystemTest', function () {
    requireRole_(['ADMIN']);
    var result = runFullSystemTest_();
    PropertiesService.getScriptProperties().setProperty('DM_LAST_QA_RUN', JSON.stringify(result.createdRecordIds));
    writeAudit_(currentEmail_(), currentRole_(), 'QA_TEST_RUN', 'SYSTEM', '', '', result.passCount + '/' + result.totalSteps, 'Full system test executed');
    return successResponse_({ steps: result.steps, passCount: result.passCount, failCount: result.failCount, totalSteps: result.totalSteps }, 'Test run complete: ' + result.passCount + '/' + result.totalSteps + ' steps passed.');
  });
}

function adminCleanupTestData() {
  return safeInvoke_('adminCleanupTestData', function () {
    requireRole_(['ADMIN']);
    var raw = PropertiesService.getScriptProperties().getProperty('DM_LAST_QA_RUN');
    if (!raw) return successResponse_({ deleted: 0 }, 'No test run found to clean up.');
    var created = JSON.parse(raw);
    var deleted = 0;
    var idFieldMap = {
      PRODUCTS: 'ProductId', SUPPLIERS: 'SupplierId', PURCHASE_ORDERS: 'PoId', RECEIVING: 'ReceivingId',
      LIVE_BATCHES: 'BatchId', QC_RECORDS: 'QcId', ORDERS: 'OrderId', ORDER_LINES: 'LineId', PAYMENTS: 'PaymentId',
      INVOICES: 'InvoiceId', PROCESSING_BATCHES: 'ProcBatchId', DRESSED_INVENTORY: 'Id', CUSTOMERS: 'CustomerId',
      CUSTOMER_ADDRESSES: 'AddressId', COUPONS: 'CouponId', GIFT_CARDS: 'CardId', SUBSCRIPTIONS: 'SubscriptionId',
      RETURNS: 'ReturnId', SUPPORT_TICKETS: 'TicketId', TICKET_MESSAGES: 'MessageId', WALLET_TRANSACTIONS: 'TxnId',
      LOYALTY: 'CustomerId', REFERRALS: 'ReferralId', DELIVERIES: 'DeliveryId'
    };
    var order = ['DELIVERIES', 'TICKET_MESSAGES', 'SUPPORT_TICKETS', 'RETURNS', 'SUBSCRIPTIONS', 'GIFT_CARDS', 'COUPONS',
      'WALLET_TRANSACTIONS', 'LOYALTY', 'REFERRALS', 'INVOICES', 'PAYMENTS', 'ORDER_LINES', 'ORDERS',
      'DRESSED_INVENTORY', 'PROCESSING_BATCHES', 'QC_RECORDS', 'LIVE_BATCHES', 'RECEIVING', 'PURCHASE_ORDERS',
      'CUSTOMER_ADDRESSES', 'CUSTOMERS', 'SUPPLIERS', 'PRODUCTS'];
    order.forEach(function (sheetName) {
      (created[sheetName] || []).forEach(function (id) {
        try { if (deleteRow_(sheetName, idFieldMap[sheetName], id)) deleted++; } catch (e) { Logger.log('Cleanup failed for ' + sheetName + ' ' + id + ': ' + e); }
      });
    });
    recalcLiveInventorySummary_();
    PropertiesService.getScriptProperties().deleteProperty('DM_LAST_QA_RUN');
    writeAudit_(currentEmail_(), currentRole_(), 'QA_TEST_CLEANUP', 'SYSTEM', '', '', deleted, 'Test data cleaned up');
    return successResponse_({ deleted: deleted }, 'Cleaned up ' + deleted + ' test records.');
  });
}

/* ============================================================================
 * SECTION 32 — AUTOMATION ENGINE & TRIGGERS
 * ========================================================================== */

function runDailyAutomation() {
  try {
    checkLowInventory();
    generateDailySummary();
    checkPendingPayments();
    processDueSubscriptions();
    checkSupportTicketSla();
    writeAudit_('SYSTEM', 'SYSTEM', 'AUTOMATION_RUN', 'SYSTEM', '', '', '', 'runDailyAutomation executed');
  } catch (e) {
    Logger.log('runDailyAutomation error: ' + e);
  }
}

function runHourlyAutomation() {
  try {
    checkPendingPayments();
    updateDeliveryStatus();
    checkSupportTicketSla();
    writeAudit_('SYSTEM', 'SYSTEM', 'AUTOMATION_RUN', 'SYSTEM', '', '', '', 'runHourlyAutomation executed');
  } catch (e) {
    Logger.log('runHourlyAutomation error: ' + e);
  }
}

function checkPendingPayments() {
  var payments = sheetToObjects_('PAYMENTS').filter(function (p) { return p.VerificationStatus === 'PENDING'; });
  var stale = payments.filter(function (p) {
    var hoursOld = (new Date() - new Date(p.Timestamp)) / 3600000;
    return hoursOld > 2;
  });
  if (stale.length > 0) {
    notify_('PAYMENT_SUBMITTED_ADMIN', getConfigValue_('supportEmail', DEFAULT_CONFIG.supportEmail), { orderId: stale.map(function (p) { return p.OrderId; }).join(', '), upiRef: 'Multiple pending > 2h' });
  }
  return { staleCount: stale.length };
}

function sendOrderNotifications() {
  var orders = sheetToObjects_('ORDERS').filter(function (o) { return o.Status === 'DELIVERED'; });
  var customers = sheetToObjects_('CUSTOMERS');
  var customerMap = {};
  customers.forEach(function (c) { customerMap[c.CustomerId] = c; });
  var sent = 0;
  orders.forEach(function (o) {
    var hoursOld = (new Date() - new Date(o.UpdatedAt)) / 3600000;
    if (hoursOld >= 24) {
      transitionOrderStatus_(o.OrderId, 'COMPLETED', 'Auto-completed 24h after delivery');
      awardOrderCompletionBenefits_(o.OrderId);
      var customer = customerMap[o.CustomerId];
      if (customer) { notify_('REVIEW_REQUEST', customer.Email, { orderId: o.OrderId }); sent++; }
    }
  });
  return { notified: sent };
}

function initializeAutomation() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  var handlerNames = existingTriggers.map(function (t) { return t.getHandlerFunction(); });
  var created = [];

  if (handlerNames.indexOf('runDailyAutomation') === -1) {
    ScriptApp.newTrigger('runDailyAutomation').timeBased().everyDays(1).atHour(6).create();
    created.push('runDailyAutomation (daily @ 6am)');
  }
  if (handlerNames.indexOf('runHourlyAutomation') === -1) {
    ScriptApp.newTrigger('runHourlyAutomation').timeBased().everyHours(1).create();
    created.push('runHourlyAutomation (hourly)');
  }
  if (handlerNames.indexOf('sendOrderNotifications') === -1) {
    ScriptApp.newTrigger('sendOrderNotifications').timeBased().everyHours(6).create();
    created.push('sendOrderNotifications (every 6h)');
  }
  if (handlerNames.indexOf('processDueSubscriptions') === -1) {
    ScriptApp.newTrigger('processDueSubscriptions').timeBased().everyDays(1).atHour(7).create();
    created.push('processDueSubscriptions (daily @ 7am)');
  }
  PropertiesService.getScriptProperties().setProperty(PROP_KEYS.TRIGGERS_INITIALIZED, 'true');
  writeAudit_(currentEmail_(), currentRole_(), 'AUTOMATION_INITIALIZED', 'SYSTEM', '', '', created.join('; '), 'Triggers verified/created');
  return { created: created, totalTriggers: ScriptApp.getProjectTriggers().length };
}

function adminInitializeAutomation() {
  return safeInvoke_('adminInitializeAutomation', function () {
    requireRole_(['ADMIN']);
    return successResponse_(initializeAutomation(), 'Automation triggers verified.');
  });
}

/* ============================================================================
 * SECTION 33 — WEB APP ENTRY POINT
 * ========================================================================== */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.initialParam = e && e.parameter ? JSON.stringify(e.parameter) : '{}';
  return template.evaluate()
    .setTitle('DesiMurga™ — Farm → Fresh → Your Home')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================================
 * SECTION 34 — MASTER DIAGNOSTICS & DEBUG ENGINE
 * ========================================================================== */

function runAdminDiagnosticTest() {
  return safeInvoke_('runAdminDiagnosticTest', function () {
    requireRole_(['ADMIN', 'MANAGER']);
    var report = [];
    
    function testStep(moduleName, fn) {
      try {
        var res = fn();
        var isSuccess = res && res.success !== false;
        var detail = 'OK';
        if (res && res.data) {
          detail = Array.isArray(res.data) ? res.data.length + ' records' : 'Valid object';
        } else if (res && res.message) {
          detail = res.message;
        }
        var statusStr = isSuccess ? 'PASS' : 'FAIL';
        report.push({ module: moduleName, status: statusStr, detail: detail });
      } catch (err) {
        var errMsg = String(err.message || err);
        report.push({ module: moduleName, status: 'FAIL', detail: errMsg });
      }
    }

    testStep('Dashboard Data', function() { return getAdminDashboardData(); });
    testStep('Orders List', function() { return getAdminOrders(null); });
    testStep('Products Catalog', function() { return getAdminProducts(); });
    testStep('Payments Desk', function() { return getAdminPayments(null); });
    testStep('Reviews Moderation', function() { return getAdminReviews(); });
    testStep('Coupons Ledger', function() { return getAdminCoupons(); });
    testStep('Live Inventory', function() { return getLiveInventorySummary(); });
    testStep('Receiving History', function() { return getAdminReceivingHistory(); });
    testStep('Processing Batches', function() { return getAdminProcessingBatches(); });
    testStep('Suppliers List', function() { return getAdminSuppliers(); });
    testStep('Purchase Orders', function() { return getAdminPurchaseOrders(); });
    testStep('Deliveries Desk', function() { return getAdminDeliveries(); });
    testStep('Expenses Ledger', function() { return getAdminExpenses(); });
    testStep('Wholesale Accounts', function() { return getAdminWholesaleAccounts(); });
    testStep('Subscriptions List', function() { return getAdminSubscriptions(); });
    testStep('Returns Requests', function() { return getAdminReturns(); });
    testStep('Support Desk', function() { return getAdminTickets(null); });
    testStep('Campaigns Engine', function() { return getAdminCampaigns(); });
    testStep('Gift Cards Ledger', function() { return getAdminGiftCards(); });
    testStep('Vendor Payments', function() { return getAdminVendorPayments(); });
    testStep('CRM & LTV Analytics', function() { return getAdminCustomerAnalytics(); });
    testStep('Static Content Pages', function() { return getAllStaticPagesAdmin(); });
    testStep('Sales Report', function() { return getSalesReport(); });
    testStep('P&L Report', function() { return getProfitLossStatement(null, null); });
    testStep('Users Management', function() { return getAdminUsers(); });
    testStep('Audit Log', function() { return getAdminAuditLog(); });
    testStep('Database Health', function() { var h = runDatabaseHealthCheck(); return { success: h.healthy, message: h.healthy ? 'Database healthy' : h.issues.join(', ') }; });

    return successResponse_(report, 'Admin diagnostic suite completed.');
  });
}


/**
 * Update customer profile details and profile picture (Base64 file upload or URL)
 */
function updateCustomerProfile(payload) {
  return safeInvoke_('updateCustomerProfile', function () {
    if (!payload || !payload.mobile) {
      return errorResponse_('Mobile number is required.', 'VALIDATION_ERROR');
    }
    
    // Auto-heal: Ensure PhotoUrl header exists in Google Sheet
    var sheet = getSheet_('CUSTOMERS');
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    if (headers.indexOf('PhotoUrl') === -1) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue('PhotoUrl').setFontWeight('bold');
    }

    var cleanSearchMobile = String(payload.mobile).replace(/\D/g, '').trim();
    var customers = sheetToObjects_('CUSTOMERS');
    var customer = null;
    
    for (var i = 0; i < customers.length; i++) {
      var cMobile = String(customers[i].Mobile || '').replace(/\D/g, '').trim();
      if (cMobile === cleanSearchMobile || cMobile.endsWith(cleanSearchMobile) || cleanSearchMobile.endsWith(cMobile)) {
        customer = customers[i];
        break;
      }
    }
    
    var customerId = customer ? customer.CustomerId : upsertCustomer_(payload.name || 'Customer', payload.mobile, payload.email);
    var photoUrl = safeString_(payload.photoUrl);
    
    // Handle Direct Base64 File Upload to Google Drive
    if (payload.photoBase64 && payload.photoBase64.indexOf('base64,') !== -1) {
      try {
        var folderId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.DOCUMENTS_FOLDER_ID) || 
                       PropertiesService.getScriptProperties().getProperty(PROP_KEYS.ROOT_FOLDER_ID);
        var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
        
        var parts = payload.photoBase64.split('base64,');
        var contentType = payload.photoContentType || 'image/jpeg';
        var decoded = Utilities.base64Decode(parts[1]);
        var blob = Utilities.newBlob(decoded, contentType, 'Avatar_' + customerId + '_' + new Date().getTime() + '.jpg');
        
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Use direct Google LH3 image CDN URL for instant display
        photoUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      } catch (uploadErr) {
        Logger.log('Profile photo upload error: ' + uploadErr);
      }
    } else if (photoUrl) {
      var fileId = extractDriveId_(photoUrl);
      if (fileId) {
        photoUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
      }
    }

    var fieldsToUpdate = {};
    if (payload.name) fieldsToUpdate.Name = safeString_(payload.name);
    if (payload.email !== undefined) fieldsToUpdate.Email = safeString_(payload.email);
    if (photoUrl) fieldsToUpdate.PhotoUrl = photoUrl;
    
    updateRowFields_('CUSTOMERS', 'CustomerId', customerId, fieldsToUpdate);

    // Update or add address if provided
    if (payload.addressLine && payload.pin) {
      upsertAddress_(customerId, {
        label: 'Home',
        addressLine: payload.addressLine,
        area: payload.area || '',
        pin: payload.pin,
        lat: payload.lat || '',
        lng: payload.lng || ''
      });
    }

    writeAudit_(payload.mobile, 'CUSTOMER', 'PROFILE_UPDATED', 'CUSTOMERS', customerId, '', JSON.stringify(fieldsToUpdate), 'Customer updated profile & photo');

    return successResponse_({
      customerId: customerId,
      name: payload.name,
      email: payload.email,
      photoUrl: photoUrl
    }, 'Profile updated successfully.');
  });
}
