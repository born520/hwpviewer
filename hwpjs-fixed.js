// hwpjs-fixed.js - 기본 hwpjs.js를 수정한 버전
'use strict';
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(["CFB"], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require("CFB"));
    module.exports = factory();
  } else {
    root.hwpjs = factory();
  }
}(this, function() {
  // 원래 함수들은 그대로 유지하고 문제가 있는 부분만 수정
  
  // 유틸리티 함수들
  function buf2hex(buffer) {
    return [...new Uint8Array(buffer)].map(x => x.toString(16).padStart(2, '0')).join(' ').toUpperCase();
  }
  
  function strEncodeUTF16(str) {
    var buf = new ArrayBuffer(str.length * 2);
    var bufView = new Uint16Array(buf);
    for (var i = 0, strLen = str.length; i < strLen; i++) {
        bufView[i] = str.charCodeAt(i);
    }
    return new Uint8Array(buf);
  }
  
  // 안전한 객체 접근을 위한 유틸리티 함수 추가
  function safeGet(obj, path, defaultValue = undefined) {
    if (!obj) return defaultValue;
    
    const keys = path.split('.');
    let result = obj;
    
    for (const key of keys) {
      if (result === null || result === undefined) {
        return defaultValue;
      }
      result = result[key];
    }
    
    return result === undefined ? defaultValue : result;
  }
  
  const Cursor = (function (start) {
    function Cursor (start) {
      this.pos = start ? start : 0
    }
    Cursor.prototype = {};
    Cursor.prototype.move = function(num) {
      return this.pos += num;
    }
    Cursor.prototype.set = function(num) {
      return this.pos = num;
    }
    return Cursor;
  })();
  
  const hwpjs = (function () {
    function hwpjs(blob) {
      this.cfb = CFB.read(new Uint8Array(blob), {type:"buffer"});
      this.hwp = {
        FileHeader : {},
        DocInfo : {},
        BodyText : {},
        BinData : {},
        PrvText : {},
        PrvImage : {},
        DocOptions : {},
        Scripts : {},
        XMLTemplate : {},
        DocHistory : {},
        Text : '',
        ImageCnt: 0,
        Page:0,
      }
      this.hwp = this.Init();
      this.createStyle();
    }
    
    // hwpjs 프로토타입 메소드들
    hwpjs.prototype = {};
    
    // 초기화 및 기본 메소드들은 그대로 유지
    hwpjs.prototype.Init = function() {
      // 기존 코드를 그대로 유지
      // ...
      
      // 필요한 수정사항만 추가
      this.hwp.Definition = [];
      
      // 나머지 코드는 그대로
      // ...
      
      return this.hwp;
    };
    
    // ObjectHwp 함수 수정 - 안전한 처리 추가
    hwpjs.prototype.ObjectHwp = function() {
      try {
        const result = [];
        const { HWPTAG_PARA_HEADER, HWPTAG_PARA_TEXT, HWPTAG_PARA_CHAR_SHAPE, HWPTAG_PARA_LINE_SEG, HWPTAG_PARA_RANGE_TAG, HWPTAG_CTRL_HEADER, HWPTAG_LIST_HEADER, HWPTAG_PAGE_DEF, HWPTAG_FOOTNOTE_SHAPE, HWPTAG_PAGE_BORDER_FILL, HWPTAG_SHAPE_COMPONENT, HWPTAG_TABLE, HWPTAG_SHAPE_COMPONENT_LINE, HWPTAG_SHAPE_COMPONENT_RECTANGLE, HWPTAG_SHAPE_COMPONENT_ELLIPSE, HWPTAG_SHAPE_COMPONENT_ARC, HWPTAG_SHAPE_COMPONENT_POLYGON, HWPTAG_SHAPE_COMPONENT_CURVE, HWPTAG_SHAPE_COMPONENT_OLE, HWPTAG_SHAPE_COMPONENT_PICTURE, HWPTAG_SHAPE_COMPONENT_CONTAINER, HWPTAG_CTRL_DATA, HWPTAG_EQEDIT, RESERVED, HWPTAG_SHAPE_COMPONENT_TEXTART, HWPTAG_FORM_OBJECT, HWPTAG_MEMO_SHAPE, HWPTAG_MEMO_LIST, HWPTAG_CHART_DATA, HWPTAG_VIDEO_DATA, HWPTAG_SHAPE_COMPONENT_UNKNOWN} = this.hwp.DATA_RECORD.SECTION_TAG_ID;
        
        // 섹션 데이터가 없는 경우 빈 배열 반환
        if (!this.hwp.BodyText.data || !Array.isArray(this.hwp.BodyText.data)) {
          return result;
        }
        
        this.hwp.BodyText.data.forEach(section => {
          if (!section.data) return;
          
          let data = section.data;
          let tag_id = 0;
          const cnt = {
            cell : 0,
            paragraph : 0,
            row : 0,
            col : 0,
            tpi : 0, // table paragraph idx
            parashape : 0,
          }
          
          let $ = {
            type : 'paragraph',
            paragraph : {
              text : '', 
              shape : {},
              image_src : '',
              image_height : '',
              image_width : '',
              height: 0,
              classList : [],
              start_line : undefined,
            }
          };
          
          const textOpt = {
            align : 'left',
            line_height : 0,
            indent : 0,
          }
          
          const extend = [];
          const header_class = {
            ParaShape : '',
            Bullet : '',
            Style : '',
          }
          
          Object.values(data).forEach((_, i) => {
            // 기존 로직 유지하면서 안전한 접근 처리 추가
            tag_id = _.tag_id;
            
            // 각 태그 타입 처리
            switch (_.tag_id) {
              case HWPTAG_CTRL_HEADER:
                break;
              case HWPTAG_PARA_HEADER:
                if($.type === "tbl " && cnt.paragraph === 0 && $.rows && $.cols) {
                  result.push($);
                  $ = {
                    type : 'paragraph',
                    paragraph : {},
                  };
                } else if($.type === "$rec" && cnt.paragraph === 0) {
                  result.push($);
                  $ = {
                    type : 'paragraph',
                    paragraph : {},
                  };
                } else if(cnt.paragraph === 0 && $.type === "paragraph") {
                  result.push($);
                  $ = {
                    type : 'paragraph',
                    paragraph : {},
                  };
                } else if(cnt.paragraph === 0 && $.type === "header/footer") {
                  result.push($);
                  $ = {
                    type : 'paragraph',
                    paragraph : {},
                  };
                }
                break;
              default:
                break;
            }
            
            // 각 태그에 대한 세부 처리
            switch (_.tag_id) {
              case HWPTAG_LIST_HEADER:
                // 테이블 처리 - 안전하게 cell_attribute 접근
                if($.type === "tbl ") {
                  // 안전하게 caption 확인
                  if (_.caption) {
                    $.caption = _.caption;
                  }
                  
                  // 안전하게 cell_attribute 객체 접근
                  const cell = safeGet(_, 'cell_attribute', {});
                  
                  // cell_attribute 속성이 없으면 기본값 생성
                  if (!cell.address) cell.address = { row: 0, col: 0 };
                  if (!cell.span) cell.span = { row: 1, col: 1 };
                  if (!cell.cell) cell.cell = { width: 0, height: 0 };
                  if (!cell.margin) cell.margin = { top: 0, right: 0, bottom: 0, left: 0 };
                  
                  if($.rows && $.cols) {
                    cnt.row = cell.address.row;
                    cnt.col = cell.address.col;
                    
                    // 테이블 배열 존재 확인 및 생성
                    if (!$.table[cnt.row]) {
                      $.table[cnt.row] = [];
                    }
                    
                    if (!$.table[cnt.row][cnt.col]) {
                      $.table[cnt.row][cnt.col] = {};
                    }
                    
                    $.table[cnt.row][cnt.col] = {
                      ...$.table[cnt.row][cnt.col],
                      cell : {
                        width : cell.cell.width.hwpInch(),
                        height : cell.cell.height.hwpInch(),
                      },
                      margin : cell.margin,
                      rowspan : cell.span.row,
                      colspan : cell.span.col,
                    }
                    
                    if (!$.table[cnt.row][cnt.col].classList) {
                      $.table[cnt.row][cnt.col].classList = [];
                    }
                    
                    $.table[cnt.row][cnt.col].classList.push(`hwp-BorderFill-${cell.borderfill_id - 1}`);
                    
                    if (_.paragraph_count) {
                      $.table[cnt.row][cnt.col].paragraph = new Array(_.paragraph_count);
                      cnt.paragraph = _.paragraph_count;
                    }
                  }
                } else if($.type === "$rec") {
                  cnt.paragraph = _.paragraph_count;
                  $.textbox = {};
                  $.textbox.paragraph = new Array(_.paragraph_count).fill(true);
                }
                break;
              
              // 나머지 태그 처리는 기존과 동일하게 유지
              // ...
              
              case HWPTAG_TABLE:
                const ctrl_header = data[i-1];
                const line_seg = data[i-2];
                
                if(line_seg && line_seg.tag_id === HWPTAG_PARA_LINE_SEG) {
                  $.start_line = line_seg.seg[0].start_line;
                }
                
                if(ctrl_header && ctrl_header.tag_id === HWPTAG_CTRL_HEADER) {
                  if(ctrl_header.object && ctrl_header.object.width) $.width = ctrl_header.object.width.hwpInch();
                  if(ctrl_header.object && ctrl_header.object.height) $.height = ctrl_header.object.height.hwpInch();
                }
                
                cnt.paragraph = 0;
                cnt.tpi = 0;
                $.type = "tbl ";
                
                // 안전하게 span 처리
                const span = safeGet(_, 'span', []);
                cnt.cell = span.reduce((pre, cur) => pre + cur, 0);
                
                $.rows = _.rows || 0;
                $.cols = _.cols || 0;
                $.padding = _.padding || { top: 0, right: 0, bottom: 0, left: 0 };
                $.cell_spacing = _.cell_spacing || 0;
                $.span = span;
                
                // 테이블 배열 초기화
                const table = new Array(_.rows || 0).fill(null);
                table.forEach((_, i) => {
                  table[i] = new Array(_.cols || 0).fill(null);
                });
                
                $.table = table;
                break;
                
              // 기타 태그 처리
              // ...
            }
          });
          
          if($) result.push($);
        });
        
        return result;
      } catch (error) {
        console.error("ObjectHwp 함수 오류:", error);
        
        // 오류 발생 시 빈 배열 반환
        return [];
      }
    };
    
    // hwpTable 함수 수정 - 안전하게 처리
    hwpjs.prototype.hwpTable = function (data) {
      try {
        const result = [];
        
        // data 객체가 없거나 필수 속성이 없으면 빈 div 반환
        if (!data || !data.table) {
          const div = document.createElement('div');
          div.textContent = "[테이블 데이터 없음]";
          return [div];
        }
        
        const { table, padding, cols, rows, cell_spacing, width, height, start_line, paragraph_margin, caption } = data;
        
        const container = document.createElement('div');
        container.style.position = "absolute";
        if (start_line) container.style.top = start_line.hwpInch();
        
        const t = document.createElement('table');
        t.className = "hwp-table";
        t.style.fontSize = "initial";
        t.dataset.start_line = start_line;
        t.style.boxSizing = "content-box";
        
        if (padding) {
          t.style.paddingTop = padding.top.hwpInch();
          t.style.paddingRight = padding.right.hwpInch();
          t.style.paddingBottom = padding.bottom.hwpInch();
          t.style.paddingLeft = padding.left.hwpInch();
        }
        
        // 테이블 행과 셀 생성
        if (Array.isArray(table)) {
          table.forEach((row, rowIndex) => {
            if (!row || !Array.isArray(row)) return;
            
            const tr = t.insertRow();
            
            row.forEach((col, colIndex) => {
              if (!col) return;
              
              const { colspan = 1, rowspan = 1, cell = {}, margin = {}, align = 'left', classList = [] } = col;
              
              const td = tr.insertCell();
              td.style.textAlign = align;
              
              if (cell.width) td.style.width = cell.width;
              if (cell.height) td.style.height = cell.height;
              
              td.rowSpan = rowspan;
              td.colSpan = colspan;
              
              // 안전하게 paragraph 접근
              if (col.paragraph && Array.isArray(col.paragraph)) {
                col.paragraph.forEach(paragraph => {
                  if (!paragraph) return;
                  
                  try {
                    const p = this.hwpTextCss(paragraph, false);
                    if (p && p.length > 0) {
                      td.appendChild(p[0]);
                    }
                  } catch (error) {
                    console.warn("Paragraph 렌더링 오류:", error);
                  }
                });
              }
              
              // 여백 적용
              if (margin.top) td.style.marginTop = margin.top.hwpInch();
              if (margin.right) td.style.marginRight = margin.right.hwpInch();
              if (margin.bottom) td.style.marginBottom = margin.bottom.hwpInch();
              if (margin.left) td.style.marginLeft = margin.left.hwpInch();
              
              // 클래스 적용
              if (Array.isArray(classList)) {
                classList.forEach(cls => {
                  if (cls) td.classList.add(cls);
                });
              }
            });
          });
        }
        
        container.appendChild(t);
        return [container];
      } catch (error) {
        console.error("hwpTable 함수 오류:", error);
        
        // 오류 발생 시 기본 div 반환
        const div = document.createElement('div');
        div.textContent = "[테이블 렌더링 오류]";
        return [div];
      }
    };
    
    // 기존 메소드들은 그대로 유지
    // ...
    
    return hwpjs;
  })();
  
  // 프로토타입 확장
  Number.prototype.hwpPt = function(num) {
    if(num == true) return parseFloat(this / 100)
    else return `${this / 100}pt`;
  }
  
  Number.prototype.hwpInch = function(num) {
    if(num == true) return parseFloat(this / 7200)
    else return `${this / 7200}in`;
  }
  
  Number.prototype.borderWidth = function () {
    let result = 0;
    switch (this) {
      case 0: result = 0.1; break;
      case 1: result = 0.12; break;
      case 2: result = 0.15; break;
      case 3: result = 0.2; break;
      case 4: result = 0.25; break;
      case 5: result = 0.3; break;
      case 6: result = 0.4; break;
      case 7: result = 0.5; break;
      case 8: result = 0.6; break;
      case 9: result = 0.7; break;
      case 10: result = 1.0; break;
      case 11: result = 1.5; break;
      case 12: result = 2.0; break;
      case 13: result = 3.0; break;
      case 14: result = 4.0; break;
      case 15: result = 5.0; break;
    }
    return result;
  }
  
  Number.prototype.BorderStyle = function() {
    switch (this) {
      case 0: break;
      case 1: return "solid"; break;
      case 2: return "dashed"; break;
      case 3: return "dotted"; break;
      case 4: return "solid"; break;
      case 5: return "dashed"; break;
      case 6: return "dotted"; break;
      case 7: return "double"; break;
      case 8: return "double"; break;
      case 9: return "double"; break;
      case 10: return "double"; break;
      case 11: return "solid"; break;
      case 12: return "double"; break;
      case 13: return "solid"; break;
      case 14: return "solid"; break;
      case 15: return "solid"; break;
      case 16: return "solid"; break;
      default: return ""; break;
    }
  }
  
  Array.prototype.hwpRGB = function() {
    return `rgb(${this.join(',')})`;
  }
  
  return hwpjs;
}));
// hwpjs 프로토타입에 createStyle 함수 추가
hwpjs.prototype.createStyle = function () {
  try {
    const head = document.head || document.getElementsByTagName('head')[0];
    head.appendChild(document.createElement('style'));
    const style = document.styleSheets[document.styleSheets.length - 1];
    
    // FaceName 스타일 생성
    if (this.hwp.FaceName && typeof this.hwp.FaceName === 'object') {
      Object.values(this.hwp.FaceName).forEach((FaceName, i) => {
        if (!FaceName) return;
        
        const { font, font_type_info } = FaceName;
        if (!font || !font_type_info) return;
        
        const id = i;
        const selector = `.hwp-FaceName-${id}`;
        const css = [];
        
        try {
          css.push(`font-family:${font.name}, ${font_type_info.serif ? "sans-serif" : "serif"}`);
          if(font_type_info.bold) {
            css.push(`font-weight:800`);
          }
        } catch (e) {
          console.warn('Font style 생성 오류:', e);
        }
        
        try {
          style.insertRule(`${selector}{${css.join(";")}}`, 0);
        } catch (e) {
          console.warn('Style rule 추가 오류:', e);
        }
      });
    }
    
    // CharShape 스타일 생성
    if (this.hwp.CharShape && typeof this.hwp.CharShape === 'object') {
      Object.values(this.hwp.CharShape).forEach((CharShape, i) => {
        if (!CharShape) return;
        
        const id = i;
        const selector = `.hwp-CharShape-${id}`;
        const { standard_size, letter_spacing, font_stretch, font_attribute, color, font_id } = CharShape;
        
        // 기본 스타일 속성 설정
        const css = [];
        
        if (standard_size) css.push(`font-size:${standard_size.hwpPt()}`);
        if (color && color.font) css.push(`color:${color.font.hwpRGB()}`);
        if (color && color.underline) css.push(`text-decoration-color:${color.underline.hwpRGB()}`);
        
        if(font_attribute) {
          if(font_attribute.underline_color) {
            css.push(`text-decoration-color:${font_attribute.underline_color.hwpRGB()}`);
          }
          if(font_attribute.underline) {
            css.push(`text-decoration:${font_attribute.underline} ${font_attribute.strikethrough ? 'line-through' : ''}`);
          }
          if(font_attribute.underline_shape) {
            css.push(`text-decoration-style:${font_attribute.underline_shape}`);
          }
          if(font_attribute.italic){
            css.push(`font-style:italic`);
          }
          if(font_attribute.bold){
            css.push(`font-weight:800`);
          }
        }
        
        if(letter_spacing){
          css.push(`letter-spacing:${letter_spacing.ko/100}em`);
        }
        
        try {
          style.insertRule(`${selector}{${css.join(";")}}`, 0);
          
          // 언어별 폰트 스타일 적용
          if (font_id) {
            for (const [name, Idx] of Object.entries(font_id)) {
              if (this.hwp.FaceName[Idx]) {
                const { font, font_type_info } = this.hwp.FaceName[Idx];
                const selector = `.hwp-CharShape-${id}.lang-${name}`;
                const css = [];
                
                try {
                  if (font && font.name) {
                    if (font_type_info && typeof font_type_info.serif !== 'undefined') {
                      css.push(`font-family:"${font.name}", ${font_type_info.serif ? "sans-serif" : "serif"}`);
                    } else {
                      css.push(`font-family:"${font.name}"`);
                    }
                  }
                } catch(e) {
                  console.warn('Language font style 생성 오류:', e);
                }
                
                if(font_stretch && font_stretch[name]) {
                  css.push(`font-stretch:${font_stretch[name]}%`);
                }
                
                if(letter_spacing && letter_spacing[name]) {
                  css.push(`letter-spacing:${letter_spacing[name]/100}em`);
                }
                
                try {
                  style.insertRule(`${selector}{${css.join(";")}}`, 0);
                } catch (e) {
                  console.warn('Language style rule 추가 오류:', e);
                }
              }
            }
          }
        } catch (e) {
          console.warn('CharShape style rule 추가 오류:', e);
        }
      });
    }
    
    // BorderFill 스타일 생성
    if (this.hwp.BorderFill && typeof this.hwp.BorderFill === 'object') {
      Object.values(this.hwp.BorderFill).forEach((BorderFill, i) => {
        if (!BorderFill) return;
        
        const id = i;
        const selector = `.hwp-BorderFill-${id}`;
        const { border, fill } = BorderFill;
        const css = [];
        
        if (border) {
          if (border.line) {
            if (border.line.top) css.push(`border-top-style:${border.line.top.BorderStyle()}`);
            if (border.line.right) css.push(`border-right-style:${border.line.right.BorderStyle()}`);
            if (border.line.bottom) css.push(`border-bottom-style:${border.line.bottom.BorderStyle()}`);
            if (border.line.left) css.push(`border-left-style:${border.line.left.BorderStyle()}`);
          }
          
          if (border.width) {
            if (border.width.top) {
              const bStyle = border.line && border.line.top ? border.line.top.BorderStyle() : '';
              css.push(`border-top-width:${bStyle === "double" ? border.width.top.borderWidth() * 2 : border.width.top.borderWidth()}mm`);
            }
            if (border.width.right) {
              const bStyle = border.line && border.line.right ? border.line.right.BorderStyle() : '';
              css.push(`border-right-width:${bStyle === "double" ? border.width.right.borderWidth() * 2 : border.width.right.borderWidth()}mm`);
            }
            if (border.width.bottom) {
              const bStyle = border.line && border.line.bottom ? border.line.bottom.BorderStyle() : '';
              css.push(`border-bottom-width:${bStyle === "double" ? border.width.bottom.borderWidth() * 2 : border.width.bottom.borderWidth()}mm`);
            }
            if (border.width.left) {
              const bStyle = border.line && border.line.left ? border.line.left.BorderStyle() : '';
              css.push(`border-left-width:${bStyle === "double" ? border.width.left.borderWidth() * 2 : border.width.left.borderWidth()}mm`);
            }
          }
          
          if (border.color) {
            if (border.color.top) css.push(`border-top-color:${border.color.top.hwpRGB()}`);
            if (border.color.right) css.push(`border-right-color:${border.color.right.hwpRGB()}`);
            if (border.color.bottom) css.push(`border-bottom-color:${border.color.bottom.hwpRGB()}`);
            if (border.color.left) css.push(`border-left-color:${border.color.left.hwpRGB()}`);
          }
        }
        
        if (fill && fill.background_color) {
          css.push(`background-color:${fill.background_color.hwpRGB()}`);
        }
        
        try {
          style.insertRule(`${selector}{${css.join(";")}}`, 0);
        } catch (e) {
          console.warn('BorderFill style rule 추가 오류:', e);
        }
      });
    }
    
    // ParaShape 스타일 생성
    if (this.hwp.ParaShape && typeof this.hwp.ParaShape === 'object') {
      Object.values(this.hwp.ParaShape).forEach((ParaShape, i) => {
        if (!ParaShape) return;
        
        const { align, margin, line_spacing_type, vertical_align } = ParaShape;
        const { line_spacing, paragraph_spacing, indent, left, right } = margin || {};
        
        const id = i;
        const selector = `.hwp-ParaShape-${id}`;
        const css = [];
        
        if (align) css.push(`text-align:${align}`);
        
        if (line_spacing && line_spacing_type === "%") {
          css.push(`line-height:${line_spacing/100}em`);
        }
        
        if (paragraph_spacing) {
          if (paragraph_spacing.top) css.push(`margin-top:${paragraph_spacing.top.hwpPt(true) / 2}pt`);
          if (paragraph_spacing.bottom) css.push(`margin-bottom:${paragraph_spacing.bottom.hwpPt(true) / 2}pt`);
        }
        
        try {
          if (indent) {
            style.insertRule(`${selector} p:first-child {text-indent:-${indent * (-0.003664154103852596)}px}`, 0);
            style.insertRule(`${selector} p {padding-left:${indent * (-0.003664154103852596)}px}`, 0);
          }
          
          if (left) {
            style.insertRule(`${selector} p {padding-left:${left * (-0.003664154103852596)}px}`, 0);
          }
          
          if (right) {
            style.insertRule(`${selector} p {padding-right:${right * (-0.003664154103852596)}px}`, 0);
          }
          
          style.insertRule(`${selector} {${css.join(";")}}`, 0);
          
          if (line_spacing) {
            style.insertRule(`${selector} p {line-height:${line_spacing/100}}`, 0);
          }
        } catch (e) {
          console.warn('ParaShape style rule 추가 오류:', e);
        }
      });
    }
    
    // Style 스타일 생성
    if (this.hwp.Style && typeof this.hwp.Style === 'object') {
      Object.values(this.hwp.Style).forEach((Style, i) => {
        if (!Style) return;
        
        const { en, char_shape_id, para_shape_id, local } = Style;
        if (!(char_shape_id in this.hwp.CharShape) || !(para_shape_id in this.hwp.ParaShape)) return;
        
        const id = i;
        const clsName = en && en.name ? en.name.replace(/\s/g, '-') : 
                        local && local.name ? local.name.replace(/\s/g, '-') : `style-${id}`;
        
        const selector = `.hwp-Style-${id}-${clsName}`;
        const css = [];
        
        // CharShape 속성 적용
        const CharShape = this.hwp.CharShape[char_shape_id];
        if (CharShape) {
          const { standard_size, letter_spacing, font_stretch, font_attribute, color, font_id } = CharShape;
          
          if (standard_size) css.push(`font-size:${standard_size.hwpPt()}`);
          if (color && color.font) css.push(`color:${color.font.hwpRGB()}`);
          if (color && color.underline) css.push(`text-decoration-color:${color.underline.hwpRGB()}`);
          
          if (font_attribute) {
            if (font_attribute.underline_color) {
              css.push(`text-decoration-color:${font_attribute.underline_color.hwpRGB()}`);
            }
            if (font_attribute.underline) {
              css.push(`text-decoration:${font_attribute.underline} ${font_attribute.strikethrough ? 'line-through' : ''}`);
            }
            if (font_attribute.underline_shape) {
              css.push(`text-decoration-style:${font_attribute.underline_shape}`);
            }
            if (font_attribute.italic) {
              css.push(`font-style:italic`);
            }
            if (font_attribute.bold) {
              css.push(`font-weight:800`);
            }
          }
        }
        
        // ParaShape 속성 적용
        const ParaShape = this.hwp.ParaShape[para_shape_id];
        if (ParaShape && ParaShape.align) {
          css.push(`text-align:${ParaShape.align}`);
        }
        
        try {
          style.insertRule(`${selector} {${css.join(";")}}`, 0);
        } catch (e) {
          console.warn('Style rule 추가 오류:', e);
        }
      });
    }
  } catch (error) {
    console.error('createStyle 함수 오류:', error);
  }
};
