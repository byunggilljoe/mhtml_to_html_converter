from email import parser
from pathlib import Path
import email
import chardet  # 인코딩 감지를 위한 라이브러리 추가
from bs4 import BeautifulSoup  # HTML 파싱을 위한 라이브러리 추가
import re
import string
import uuid
import os
import requests  # 웹 폰트 다운로드를 위해 추가
from urllib.parse import urljoin, urlparse  # URL 처리를 위해 추가

def parse_mhtml_file(file_path):
    # MHTML 파일 경로
    mhtml_path = Path(file_path)
    
    # 리소스 디렉토리 생성
    resource_dir = Path("resource")
    image_dir = resource_dir / "image"
    css_dir = resource_dir / "css"
    js_dir = resource_dir / "javascript"
    html_dir = resource_dir / "html"
    font_dir = resource_dir / "font"  # 폰트 디렉토리 추가
    
    # 필요한 디렉토리 생성
    for dir_path in [resource_dir, image_dir, css_dir, js_dir, html_dir, font_dir]:  # font_dir 추가
        dir_path.mkdir(exist_ok=True)
    
    # 리소스 매핑 딕셔너리
    resource_mapping = {}
    html_saved = False
    html_content = None  # HTML 컨텐츠를 저장할 변수 추가
    
    # Content-ID 매핑을 위한 딕셔너리 추가
    cid_mapping = {}
    
    # 추가 HTML 파일들을 저장할 리스트
    additional_html_files = []  # [(payload, save_path, original_path, content_id), ...]
    
    # 전체 리소스 교체 카운터
    total_replacement_count = 0
    
    def sanitize_filename(filename):
        # 쿼리 파라미터 제거 (?로 시작하는 부분)
        filename = filename.split('?')[0]
        
        # 파일 확장자 분리
        name, ext = os.path.splitext(filename)
        
        # 파일명 길이 제한 (확장자 제외)
        if len(name) > 50:  # 적절한 길이로 제한
            name = name[:50]
        
        # 유효한 파일명 문자만 허용
        valid_chars = f"-_.{string.ascii_letters}{string.digits}"
        # 유효하지 않은 문자를 '_'로 변경
        sanitized_name = ''.join(c if c in valid_chars else '_' for c in name)
        # 빈 파일명인 경우 UUID 생성
        if not sanitized_name:
            sanitized_name = str(uuid.uuid4())[:8]
        return f"{sanitized_name}{ext}"

    def download_web_font(font_url, base_url=None):
        """웹 폰트를 다운로드하고 로컬에 저장"""
        try:
            # 상대 URL인 경우 절대 URL로 변환
            if base_url and not urlparse(font_url).netloc:
                font_url = urljoin(base_url, font_url)
            
            # 이미 다운로드된 폰트인지 확인
            if font_url in resource_mapping:
                print(f"Font already downloaded: {font_url}")
                return resource_mapping[font_url]
            
            # 폰트 다운로드
            response = requests.get(font_url, allow_redirects=True)
            if response.status_code == 200:
                # 파일명 생성
                url_path = urlparse(font_url).path
                filename = os.path.basename(url_path)
                if not filename:
                    filename = str(uuid.uuid4())[:8]
                
                # 파일명 정리
                sanitized_filename = sanitize_filename(filename)
                if not any(sanitized_filename.endswith(ext) for ext in ['.woff', '.woff2', '.ttf', '.eot', '.otf']):
                    # Content-Type에서 확장자 추출 시도
                    content_type = response.headers.get('Content-Type', '')
                    if 'woff2' in content_type:
                        sanitized_filename = f"{sanitized_filename}.woff2"
                    elif 'woff' in content_type:
                        sanitized_filename = f"{sanitized_filename}.woff"
                    elif 'ttf' in content_type or 'truetype' in content_type:
                        sanitized_filename = f"{sanitized_filename}.ttf"
                    elif 'opentype' in content_type:
                        sanitized_filename = f"{sanitized_filename}.otf"
                    elif 'embedded-opentype' in content_type:
                        sanitized_filename = f"{sanitized_filename}.eot"
                    else:
                        # 확장자를 추측할 수 없는 경우 woff2 사용
                        sanitized_filename = f"{sanitized_filename}.woff2"
                
                # 폰트 파일 저장
                save_path = font_dir / sanitized_filename
                save_path.write_bytes(response.content)
                
                # 상대 경로로 변환
                relative_save_path = str(save_path)
                if os.path.sep == '\\':
                    relative_save_path = relative_save_path.replace('\\', '/')
                
                # 리소스 매핑에 추가
                resource_mapping[font_url] = relative_save_path
                print(f"Downloaded and saved font: {font_url} -> {relative_save_path}")
                return relative_save_path
            else:
                print(f"Failed to download font: {font_url} (Status code: {response.status_code})")
                return None
        except Exception as e:
            print(f"Error downloading font: {font_url} - {str(e)}")
            return None

    def save_content(part):
        nonlocal html_saved
        nonlocal html_content
        nonlocal additional_html_files  # additional_html_files를 외부 스코프에서 사용
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
            # HTML 메인 컨텐츠 나중에 처리하기 위해 저장
            if content_type == 'text/html':
                if not html_saved:
                    html_content = payload
                    html_saved = True
                    return
                else:
                    # 추가 HTML 파일들은 나중에 처리하기 위해 저장
                    filename = Path(content_location).name if content_location else ""
                    if not filename:
                        filename = str(uuid.uuid4())[:8]
                    
                    # URL에서 파일명만 추출 (경로와 쿼리 파라미터 제거)
                    filename = filename.split('/')[-1].split('?')[0]
                    
                    # 파일명 정리
                    sanitized_filename = sanitize_filename(filename)
                    if not sanitized_filename.endswith('.html'):
                        sanitized_filename = f"{sanitized_filename}.html"
                    
                    # HTML 파일 경로 설정
                    save_path = html_dir / sanitized_filename
                    
                    # 나중에 처리하기 위해 정보 저장
                    additional_html_files.append((payload, save_path, original_path, content_id))
                    
                    # 리소스 매핑 저장 (파일 경로만)
                    relative_save_path = str(save_path)
                    if os.path.sep == '\\':
                        relative_save_path = relative_save_path.replace('\\', '/')
                    
                    # 일반 경로 매핑
                    if original_path:
                        resource_mapping[original_path] = relative_save_path
                        print(f"Mapped HTML path: {original_path} -> {relative_save_path}")
                    
                    # Content-ID 매핑
                    if content_id:
                        cid_url = f"cid:{content_id}"
                        resource_mapping[cid_url] = relative_save_path
                        resource_mapping[content_id] = relative_save_path
                        cid_mapping[content_id] = relative_save_path
                        print(f"Mapped HTML CID: {cid_url} -> {relative_save_path}")
                        print(f"Mapped HTML CID (without prefix): {content_id} -> {relative_save_path}")
                    return
            
            # 리소스 파일 처리
            filename = Path(content_location).name if content_location else ""
            if not filename:
                filename = str(uuid.uuid4())[:8]  # UUID를 문자열로 변환 후 슬라이싱
            
            # URL에서 파일명만 추출 (경로와 쿼리 파라미터 제거)
            filename = filename.split('/')[-1].split('?')[0]
            
            # 파일명 정리
            sanitized_filename = sanitize_filename(filename)
            save_path = None
            
            # 리소스 타입별 저장
            if 'font' in content_type or any(filename.endswith(ext) for ext in ['.woff', '.woff2', '.ttf', '.eot', '.otf']):
                # 폰트 파일 확장자 처리
                if not any(sanitized_filename.endswith(ext) for ext in ['.woff', '.woff2', '.ttf', '.eot', '.otf']):
                    ext = f".{content_type.split('/')[-1]}"
                    if ext == '.vnd.ms-fontobject':
                        ext = '.eot'
                    sanitized_filename = f"{sanitized_filename}{ext}"
                save_path = font_dir / sanitized_filename
                save_path.write_bytes(payload)
                print(f"Saved font file: {save_path}")
                
            elif 'image' in content_type:
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
                
                # CSS 파일에서 폰트 URL 찾기
                try:
                    css_content = payload.decode('utf-8', errors='ignore')
                    # @font-face 규칙에서 src: url() 찾기
                    font_urls = re.findall(r'url\([\'"]?([^\'"]+\.(?:woff2?|ttf|eot|otf))[\'"]?\)', css_content)
                    
                    # 발견된 각 폰트 URL에 대해 처리
                    for font_url in font_urls:
                        print(f"Found font reference in CSS: {font_url}")
                        local_path = download_web_font(font_url, content_location)
                        if local_path:
                            # CSS 내용의 폰트 URL을 로컬 경로로 교체
                            css_content = css_content.replace(font_url, local_path)
                    
                    # 수정된 CSS 저장
                    save_path.write_text(css_content, encoding='utf-8')
                except Exception as e:
                    print(f"Warning: Failed to process CSS content for font detection: {e}")
                    save_path.write_bytes(payload)
                
                # CSS 파일에서 폰트 URL 찾기
                try:
                    css_content = payload.decode('utf-8', errors='ignore')
                    # @font-face 규칙에서 src: url() 찾기
                    font_urls = re.findall(r'url\([\'"]?([^\'"]+\.(?:woff2?|ttf|eot|otf))[\'"]?\)', css_content)
                    for font_url in font_urls:
                        print(f"Found font reference in CSS: {font_url}")
                        if font_url in resource_mapping:
                            print(f"Font already mapped: {font_url} -> {resource_mapping[font_url]}")
                except Exception as e:
                    print(f"Warning: Failed to process CSS content for font detection: {e}")
                
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
            exit(0)

    def process_html(payload):
        nonlocal total_replacement_count  # 전역 카운터 사용
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
        
        # 디버깅: 매핑 정보 출력
        print("\nResource mapping contents:")
        print("CID Mapping:")
        for k, v in cid_mapping.items():
            print(f"  {k} -> {v}")
        print("\nResource Mapping:")
        for k, v in resource_mapping.items():
            print(f"  {k} -> {v}")
        print("\nStarting resource replacement...")
        
        # 리소스 교체 카운터 추가
        replacement_count = 0
        unmapped_resources = set()  # 매핑되지 않은 리소스 추적
        
        # 리소스 경로 업데이트
        for tag in soup.find_all(['img', 'script', 'link', 'iframe', 'frame']):
            src_attr = 'href' if tag.name == 'link' else 'src'
            resource_path = tag.get(src_attr)
            
            if not resource_path:
                continue
                
            print(f"Processing {tag.name} with {src_attr}: {resource_path}")  # 디버깅용
            
            # cid: URL 처리
            if resource_path.startswith('cid:'):
                cid_without_prefix = resource_path[4:]  # cid: 제거
                if cid_without_prefix in cid_mapping:  # cid_mapping 사용
                    tag[src_attr] = cid_mapping[cid_without_prefix]
                    replacement_count += 1
                    print(f"Replaced ({replacement_count}) {resource_path} -> {cid_mapping[cid_without_prefix]}")
                elif cid_without_prefix in resource_mapping:
                    tag[src_attr] = resource_mapping[cid_without_prefix]
                    replacement_count += 1
                    print(f"Replaced ({replacement_count}) {resource_path} -> {resource_mapping[cid_without_prefix]}")
                elif resource_path in resource_mapping:
                    tag[src_attr] = resource_mapping[resource_path]
                    replacement_count += 1
                    print(f"Replaced ({replacement_count}) {resource_path} -> {resource_mapping[resource_path]}")
                else:
                    unmapped_resources.add(f"{tag.name}[{src_attr}]: {resource_path}")
            # 일반 경로 처리
            elif resource_path in resource_mapping:
                tag[src_attr] = resource_mapping[resource_path]
                replacement_count += 1
                print(f"Replaced ({replacement_count}) {resource_path} -> {resource_mapping[resource_path]}")
            elif not resource_path.startswith(('http://', 'https://', 'data:', '/')):
                # 외부 URL이나 절대 경로가 아닌 경우만 추적
                unmapped_resources.add(f"{tag.name}[{src_attr}]: {resource_path}")
        
        # 인라인 스타일의 url() 처리
        for tag in soup.find_all(style=True):
            style = tag['style']
            # cid: URL을 포함한 모든 URL 패턴 찾기
            urls = re.findall(r'url\([\'"]?(cid:[^\'"]+|[^\'"]+)[\'"]?\)', style)
            for url in urls:
                replaced = False
                if url.startswith('cid:'):
                    cid_without_prefix = url[4:]
                    if cid_without_prefix in cid_mapping:  # 먼저 cid_mapping 확인
                        new_url = cid_mapping[cid_without_prefix]
                        style = style.replace(url, new_url)
                        replacement_count += 1
                        replaced = True
                        print(f"Replaced ({replacement_count}) style URL {url} -> {new_url}")
                    elif cid_without_prefix in resource_mapping:
                        new_url = resource_mapping[cid_without_prefix]
                        style = style.replace(url, new_url)
                        replacement_count += 1
                        replaced = True
                        print(f"Replaced ({replacement_count}) style URL {url} -> {new_url}")
                    elif url in resource_mapping:
                        new_url = resource_mapping[url]
                        style = style.replace(url, new_url)
                        replacement_count += 1
                        replaced = True
                        print(f"Replaced ({replacement_count}) style URL {url} -> {new_url}")
                elif url in resource_mapping:
                    new_url = resource_mapping[url]
                    style = style.replace(url, new_url)
                    replacement_count += 1
                    replaced = True
                    print(f"Replaced ({replacement_count}) style URL {url} -> {new_url}")
                
                if not replaced and not url.startswith(('http://', 'https://', 'data:', '/')):
                    unmapped_resources.add(f"style[url()]: {url}")
            tag['style'] = style
        
        # 변환된 HTML 출력 (디버깅용)
        print("\nProcessed HTML preview:")
        for tag in soup.find_all(['link', 'script', 'img', 'iframe', 'frame']):
            print(f"{tag.name}: {tag.get('href') or tag.get('src')}")
        
        print(f"\nResource replacements in this file: {replacement_count}")
        total_replacement_count += replacement_count  # 전체 카운터에 추가
        
        # 매핑되지 않은 리소스 보고
        if unmapped_resources:
            print("\nUnmapped Resources Report:")
            print("The following resources could not be mapped to local files:")
            for resource in sorted(unmapped_resources):
                print(f"  - {resource}")
        
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
    
    # 모든 리소스가 처리된 후 HTML 처리
    if html_saved and html_content:
        print("Processing main HTML content...")
        process_html(html_content)
        print("Processing main HTML completed.")
        
        # 추가 HTML 파일들 처리
        print("\nProcessing additional HTML files...")
        for payload, save_path, original_path, content_id in additional_html_files:
            try:
                content = payload.decode('utf-8', errors='ignore')
                soup = BeautifulSoup(content, 'html.parser')
                
                # 이 파일의 리소스 교체 카운터
                file_replacement_count = 0
                
                # 리소스 경로 업데이트
                for tag in soup.find_all(['img', 'script', 'link', 'iframe', 'frame']):
                    src_attr = 'href' if tag.name == 'link' else 'src'
                    resource_path = tag.get(src_attr)
                    
                    if not resource_path:
                        continue
                    
                    if resource_path.startswith('cid:'):
                        cid_without_prefix = resource_path[4:]
                        if cid_without_prefix in cid_mapping:
                            tag[src_attr] = cid_mapping[cid_without_prefix]
                            file_replacement_count += 1
                        elif cid_without_prefix in resource_mapping:
                            tag[src_attr] = resource_mapping[cid_without_prefix]
                            file_replacement_count += 1
                        elif resource_path in resource_mapping:
                            tag[src_attr] = resource_mapping[resource_path]
                            file_replacement_count += 1
                    elif resource_path in resource_mapping:
                        tag[src_attr] = resource_mapping[resource_path]
                        file_replacement_count += 1
                
                # 인라인 스타일의 url() 처리
                for tag in soup.find_all(style=True):
                    style = tag['style']
                    urls = re.findall(r'url\([\'"]?(cid:[^\'"]+|[^\'"]+)[\'"]?\)', style)
                    for url in urls:
                        if url.startswith('cid:'):
                            cid_without_prefix = url[4:]
                            if cid_without_prefix in cid_mapping:
                                style = style.replace(url, cid_mapping[cid_without_prefix])
                                file_replacement_count += 1
                            elif cid_without_prefix in resource_mapping:
                                style = style.replace(url, resource_mapping[cid_without_prefix])
                                file_replacement_count += 1
                            elif url in resource_mapping:
                                style = style.replace(url, resource_mapping[url])
                                file_replacement_count += 1
                        elif url in resource_mapping:
                            style = style.replace(url, resource_mapping[url])
                            file_replacement_count += 1
                    tag['style'] = style
                
                # 수정된 HTML 저장
                save_path.write_text(str(soup), encoding='utf-8')
                print(f"Processed and saved HTML file: {save_path}")
                print(f"Resource replacements in this file: {file_replacement_count}")
                total_replacement_count += file_replacement_count  # 전체 카운터에 추가
            except Exception as e:
                print(f"Warning: Failed to process HTML content: {e}")
                # 실패하면 원본 그대로 저장
                save_path.write_bytes(payload)
                print(f"Saved original HTML content: {save_path}")
        print("Processing additional HTML files completed.")
        
        # 모든 처리가 끝난 후 총 교체 수 출력
        print(f"\nTotal number of resource replacements across all HTML files: {total_replacement_count}")
    else:
        print("No HTML content found or processing failed.")
        print(html_saved, html_content is None)

# 사용 예시
if __name__ == "__main__":
    mhtml_file = "./original.mhtml"  # MHTML 파일 경로
    parse_mhtml_file(mhtml_file)