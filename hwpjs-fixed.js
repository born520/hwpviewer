'use strict';
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(["CFB"], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require("CFB"));
  } else {
    root.hwpjs = factory();
  }
}(this, function() {
  // 기본 유틸리티 함수
  function buf2hex(buffer) {
    return [...new Uint8Array(buffer)].map(x => x.toString(16).padStart(2, '0')).join(' ').toUpperCase();
  }
  
  // 간단한 hwpjs 구현
  function hwpjs(blob) {
    try {
      this.cfb = CFB.read(new Uint8Array(blob), {type:"buffer"});
      this.hwp = {
        FileHeader: {},
        DocInfo: {},
        BodyText: {},
        BinData: {},
        Text: '',
        ImageCnt: 0
      };
      
      // 파일 구조 파싱
      this.parseFileStructure();
      
      // 텍스트 추출
      this.extractText();
    } catch (error) {
      console.error("hwpjs 초기화 오류:", error);
    }
  }
  
  // 프로토타입 메소드
  hwpjs.prototype = {
    // 파일 구조 파싱
    parseFileStructure: function() {
      try {
        // 필수 파일 인덱스 찾기
        this.hwp.FileHeader = this.cfb.FileIndex.find(FileIndex => FileIndex.name === "FileHeader");
        this.hwp.DocInfo = this.cfb.FileIndex.find(FileIndex => FileIndex.name === "DocInfo");
        this.hwp.BodyText = this.cfb.FileIndex.find(FileIndex => FileIndex.name === "BodyText");
        this.hwp.BinData = this.cfb.FileIndex.find(FileIndex => FileIndex.name === "BinData");
        
        // 파일 헤더 정보 파싱
        if (this.hwp.FileHeader && this.hwp.FileHeader.content) {
          this.hwp.FileHeader.data = {
            signature: this.textDecoder(this.hwp.FileHeader.content.slice(0, 32)),
            version: this.hwp.FileHeader.content.slice(32, 36)
          };
        }
        
        // BinData 폴더 찾기 (이미지용)
        const binDataFolders = this.cfb.FullPaths.filter(path => path.includes('BinData/'));
        this.hwp.ImageCnt = binDataFolders.length;
      } catch (error) {
        console.error("파일 구조 파싱 오류:", error);
      }
    },
    
    // 텍스트 추출
    extractText: function() {
      try {
        // BodyText에서 텍스트 추출
        if (this.hwp.BodyText && this.hwp.BodyText.content) {
          try {
            // 압축 해제 시도
            const content = pako.inflate(this.uint_8(this.hwp.BodyText.content), { windowBits: -15 });
            
            // 텍스트 추출 (매우 단순화된 버전)
            let text = '';
            for (let i = 0; i < content.length; i += 2) {
              const charCode = content[i] | (content[i + 1] << 8);
              
              // 일반 텍스트 문자인 경우만 추출
              if (charCode >= 32 && charCode <= 126 || charCode >= 0xAC00 && charCode <= 0xD7A3) {
                text += String.fromCharCode(charCode);
              } else if (charCode === 10 || charCode === 13) {
                text += '\n';
              }
            }
            
            this.hwp.Text = text;
          } catch (inflateError) {
            console.warn("압축 해제 오류:", inflateError);
            
            // 압축 해제 실패 시 직접 텍스트 추출 시도
            const content = this.hwp.BodyText.content;
            let text = '';
            
            for (let i = 0; i < content.length; i++) {
              const b = content[i];
              if (b >= 32 && b <= 126) {
                text += String.fromCharCode(b);
              }
            }
            
            this.hwp.Text = text;
          }
        }
        
        // DocInfo에서 부가 정보 추출
        if (this.hwp.DocInfo && this.hwp.DocInfo.content) {
          try {
            const content = pako.inflate(this.uint_8(this.hwp.DocInfo.content), { windowBits: -15 });
            
            // 여기서 추가 정보 추출 가능 (단순화를 위해 생략)
          } catch (error) {
            console.warn("DocInfo 파싱 오류:", error);
          }
        }
      } catch (error) {
        console.error("텍스트 추출 오류:", error);
        this.hwp.Text = '텍스트 추출 실패';
      }
    },
    
    // 텍스트 디코딩 함수
    textDecoder: function(uint_8, type = 'utf8') {
      try {
        const decoder = new TextDecoder(type);
        return decoder.decode(new Uint8Array(uint_8));
      } catch (error) {
        console.warn("텍스트 디코딩 오류:", error);
        return '';
      }
    },
    
    // Uint8Array 변환 함수
    uint_8: function(data) {
      return new Uint8Array(data);
    },
    
    // 텍스트 반환 함수
    getText: function() {
      return this.hwp.Text || '텍스트 없음';
    },
    
    // HTML 렌더링 함수 (매우 단순화)
    getHtml: function() {
      const text = this.getText();
      
      return `
        <div class="hwp-document">
          <div class="hwp-page">
            <div class="hwp-content">
              <p>${text.replace(/\n/g, '</p><p>')}</p>
            </div>
          </div>
        </div>
      `;
    },
    
    // 이미지 추출 함수
    getBinImage: function(index) {
      try {
        // BinData 폴더의 이미지 파일 찾기
        const binDataPaths = this.cfb.FullPaths.filter(path => 
          path.includes('BinData/') && 
          (path.toLowerCase().endsWith('.jpg') || 
           path.toLowerCase().endsWith('.png') || 
           path.toLowerCase().endsWith('.bmp') || 
           path.toLowerCase().endsWith('.gif'))
        );
        
        if (index < 0 || index >= binDataPaths.length) {
          console.warn("존재하지 않는 이미지 인덱스:", index);
          return null;
        }
        
        const path = binDataPaths[index];
        const fileIdx = this.cfb.FileIndex.findIndex(file => file.name === path);
        
        if (fileIdx === -1) {
          console.warn("이미지 파일을 찾을 수 없습니다:", path);
          return null;
        }
        
        const imgFile = this.cfb.FileIndex[fileIdx];
        
        // 확장자 추출
        const extMatch = path.match(/\.([^.]+)$/);
        const ext = extMatch ? extMatch[1].toLowerCase() : 'png';
        
        // 압축 해제 시도
        let imgData;
        try {
          imgData = pako.inflate(this.uint_8(imgFile.content), { windowBits: -15 });
        } catch (error) {
          console.warn("이미지 압축 해제 오류, 원본 사용:", error);
          imgData = imgFile.content;
        }
        
        // 이미지 생성
        const img = new Image();
        img.src = URL.createObjectURL(new Blob([new Uint8Array(imgData)], { type: `image/${ext}` }));
        
        return img;
      } catch (error) {
        console.error("이미지 추출 오류:", error);
        return null;
      }
    }
  };
  
  return hwpjs;
}));
