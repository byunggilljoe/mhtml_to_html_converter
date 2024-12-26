from email import parser
from pathlib import Path
import email
import chardet  # 인코딩 감지를 위한 라이브러리 추가
from bs4 import BeautifulSoup  # HTML 파싱을 위한 라이브러리 추가
import re
import string
import uuid
import os

def parse_mhtml_file(file_path):
    # MHTML 파일 경로
    mhtml_path = Path(file_path)
    
    # 리소스 디렉토리 생성
    resource_dir = Path("resource")
    image_dir = resource_dir / "image"
    css_dir = resource_dir / "css"
    js_dir = resource_dir / "javascript"
    
    # 필요한 디렉토리 생성
    for dir_path in [resource_dir, image_dir, css_dir, js_dir]:
        dir_path.mkdir(exist_ok=True)
    
    # 리소스 매핑 딕셔너리
    resource_mapping = {}
    html_saved = False
    
    # Content-ID 매핑을 위한 딕셔너리 추가
    cid_mapping = {}
    
    def sanitize_filename(filename):
        # 쿼리 파라미터 제거 (?로 시작하는 부분)
        filename = filename.split('?')[0]
        
        # 유효한 파일명 문자만 허용
        valid_chars = f"-_.{string.ascii_letters}{string.digits}"
        # 파일 확장자 분리
        name, ext = os.path.splitext(filename)
        # 유효하지 않은 문자를 '_'로 변경
        sanitized_name = ''.join(c if c in valid_chars else '_' for c in name)
        # 빈 파일명인 경우 UUID 생성
        if not sanitized_name:
            sanitized_name = str(uuid.uuid4())[:8]
        return f"{sanitized_name}{ext}"

    def save_content(part):
        nonlocal html_saved
        content_type = part.get_content_type()
        content_location = part.get("Content-Location", "")
        content_id = part.get("Content-ID", "")
        original_path = content_location
        
        # Content-ID 처리 (< > 제거)
        if content_id:
            content_id = content_id.strip("<>")
            
        payload = part.get_payload(decode=True)
        if not payload:
            return
            
        try:
            # HTML 메인 컨텐츠 처리
            if content_type == 'text/html' and not html_saved:
                html_content = process_html(payload)
                html_saved = True
                return
            
            # 리소스 파일 처리
            filename = Path(content_location).name if content_location else ""
            if not filename:
                filename = f"{uuid.uuid4()[:8]}"
            
            # URL에서 파일명만 추출 (경로와 쿼리 파라미터 제거)
            filename = filename.split('/')[-1].split('?')[0]
            
            # 파일명 정리
            sanitized_filename = sanitize_filename(filename)
            save_path = None
            
            # 리소스 타입별 저장
            if 'image' in content_type:
                if not sanitized_filename.endswith(tuple(['.jpg','.jpeg','.png','.gif','.webp','.svg'])):
                    if 'svg' in content_type:
                        sanitized_filename = f"{sanitized_filename}.svg"
                    else:
                        ext = f".{content_type.split('/')[-1]}"
                        sanitized_filename = f"{sanitized_filename}{ext}"
                save_path = image_dir / sanitized_filename
                save_path.write_bytes(payload)
                
            elif 'css' in content_type or filename.endswith('.css'):
                if not sanitized_filename.endswith('.css'):
                    sanitized_filename = f"{sanitized_filename}.css"
                save_path = css_dir / sanitized_filename
                save_path.write_text(payload.decode('utf-8', errors='ignore'))
                
            elif 'javascript' in content_type or content_location.endswith('.js'):
                if not sanitized_filename.endswith('.js'):
                    sanitized_filename = f"{sanitized_filename}.js"
                save_path = js_dir / sanitized_filename
                save_path.write_text(payload.decode('utf-8', errors='ignore'))
            
            # 리소스 매핑 저장
            if save_path:
                relative_save_path = str(save_path)
                if os.path.sep == '\\':
                    relative_save_path = relative_save_path.replace('\\', '/')
                
                # 일반 경로 매핑
                if original_path:
                    resource_mapping[original_path] = relative_save_path
                    print(f"Mapped path: {original_path} -> {relative_save_path}")
                
                # Content-ID 매핑 (cid: 제거된 버전도 함께 저장)
                if content_id:
                    cid_url = f"cid:{content_id}"
                    resource_mapping[cid_url] = relative_save_path
                    resource_mapping[content_id] = relative_save_path  # cid: 없는 버전도 매핑
                    cid_mapping[content_id] = relative_save_path
                    print(f"Mapped CID: {cid_url} -> {relative_save_path}")
                    print(f"Mapped CID (without prefix): {content_id} -> {relative_save_path}")
                    
        except Exception as e:
            print(f"Error processing content: {e}\n")
            print(f"Failed content type: {content_type}")
            print(f"Original path: {original_path}")
            print(f"Content-ID: {content_id}")

    def process_html(payload):
        # HTML 디코딩
        detected = chardet.detect(payload)
        try:
            content = payload.decode(detected['encoding'])
        except:
            for encoding in ['utf-8', 'cp949', 'euc-kr', 'iso-8859-1']:
                try:
                    content = payload.decode(encoding)
                    break
                except:
                    continue
        
        soup = BeautifulSoup(content, 'html.parser')
        
        # 리소스 경로 업데이트
        for tag in soup.find_all(['img', 'script', 'link']):
            src_attr = 'href' if tag.name == 'link' else 'src'
            resource_path = tag.get(src_attr)
            
            if not resource_path:
                continue
                
            print(f"Processing {tag.name} with {src_attr}: {resource_path}")  # 디버깅용
            
            # cid: URL 처리
            if resource_path.startswith('cid:'):
                cid_without_prefix = resource_path[4:]  # cid: 제거
                if cid_without_prefix in resource_mapping:
                    tag[src_attr] = resource_mapping[cid_without_prefix]
                    print(f"Replaced {resource_path} -> {resource_mapping[cid_without_prefix]}")
                elif resource_path in resource_mapping:
                    tag[src_attr] = resource_mapping[resource_path]
                    print(f"Replaced {resource_path} -> {resource_mapping[resource_path]}")
            # 일반 경로 처리
            elif resource_path in resource_mapping:
                tag[src_attr] = resource_mapping[resource_path]
                print(f"Replaced {resource_path} -> {resource_mapping[resource_path]}")
        
        # 인라인 스타일의 url() 처리
        for tag in soup.find_all(style=True):
            style = tag['style']
            # cid: URL을 포함한 모든 URL 패턴 찾기
            urls = re.findall(r'url\([\'"]?(cid:[^\'"]+|[^\'"]+)[\'"]?\)', style)
            for url in urls:
                if url.startswith('cid:'):
                    cid_without_prefix = url[4:]
                    if cid_without_prefix in resource_mapping:
                        new_url = resource_mapping[cid_without_prefix]
                        style = style.replace(url, new_url)
                        print(f"Replaced style URL {url} -> {new_url}")
                    elif url in resource_mapping:
                        new_url = resource_mapping[url]
                        style = style.replace(url, new_url)
                        print(f"Replaced style URL {url} -> {new_url}")
                elif url in resource_mapping:
                    new_url = resource_mapping[url]
                    style = style.replace(url, new_url)
                    print(f"Replaced style URL {url} -> {new_url}")
            tag['style'] = style
        
        # 변환된 HTML 출력 (디버깅용)
        print("\nProcessed HTML preview:")
        for tag in soup.find_all(['link', 'script', 'img']):
            print(f"{tag.name}: {tag.get('href') or tag.get('src')}")
        
        # 수정된 HTML 저장
        save_path = Path("main.html")
        save_path.write_text(str(soup), encoding='utf-8')
        return content

    # MHTML 파싱
    with open(mhtml_path, 'r', encoding='utf-8') as f:
        email_parser = parser.Parser()
        mhtml = email_parser.parse(f)
        
    # 먼저 모든 리소스를 처리하고 매핑 생성
    if mhtml.is_multipart():
        for part in mhtml.walk():
            save_content(part)
    else:
        save_content(mhtml)

# 사용 예시
if __name__ == "__main__":
    mhtml_file = "./original.mhtml"  # MHTML 파일 경로
    parse_mhtml_file(mhtml_file)