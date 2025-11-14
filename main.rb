# ============================================
# main.rb (Professor Bot - 안정화 완전판)
# ============================================
require 'bundler/setup'
require 'dotenv'
require 'time'
require 'json'
require 'ostruct'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'net/http'
require_relative 'mastodon_client'
require_relative 'sheet_manager'
require_relative 'professor_command_parser'

Dotenv.load('.env')

# 환경 변수 검증
required_envs = %w[MASTODON_DOMAIN ACCESS_TOKEN SHEET_ID GOOGLE_CREDENTIALS_PATH]
missing = required_envs.select { |v| ENV[v].nil? || ENV[v].strip.empty? }
if missing.any?
  missing.each { |v| puts "[환경변수 누락] #{v}" }
  puts "[오류] .env 파일을 확인해주세요."
  exit 1
end

DOMAIN       = ENV['MASTODON_DOMAIN']
TOKEN        = ENV['ACCESS_TOKEN']
SHEET_ID     = ENV['SHEET_ID']
CRED_PATH    = ENV['GOOGLE_CREDENTIALS_PATH']
LAST_ID_FILE = 'last_mention_id.txt'

MENTION_ENDPOINT = "https://#{DOMAIN}/api/v1/notifications"
POST_ENDPOINT    = "https://#{DOMAIN}/api/v1/statuses"

puts "[교수봇] 실행 시작 (#{Time.now.strftime('%H:%M:%S')})"

# ============================================
# Google Sheets 연결
# ============================================
begin
  creds = Google::Auth::ServiceAccountCredentials.make_creds(
    json_key_io: File.open(CRED_PATH),
    scope: ['https://www.googleapis.com/auth/spreadsheets']
  )
  creds.fetch_access_token!
  service = Google::Apis::SheetsV4::SheetsService.new
  service.authorization = creds
  sheet_manager = SheetManager.new(service, SHEET_ID)
  puts "[Google Sheets] 연결 성공"
rescue ArgumentError
  # Ruby 3.x에서 인자 해석 오류가 발생할 경우 안전 처리
  creds = Google::Auth::ServiceAccountCredentials.make_creds({json_key_io: File.open(CRED_PATH), scope: ['https://www.googleapis.com/auth/spreadsheets']})
  creds.fetch_access_token!
  service = Google::Apis::SheetsV4::SheetsService.new
  service.authorization = creds
  sheet_manager = SheetManager.new(service, SHEET_ID)
  puts "[Google Sheets] 연결 성공 (대체 방식)"
rescue => e
  puts "[에러] Google Sheets 연결 실패: #{e.message}"
  exit 1
end

# ============================================
# Mentions API 처리 함수
# ============================================
def fetch_mentions(since_id = nil)
  url = "#{MENTION_ENDPOINT}?types[]=mention&limit=20"
  url += "&since_id=#{since_id}" if since_id
  uri = URI(url)
  req = Net::HTTP::Get.new(uri)
  req['Authorization'] = "Bearer #{TOKEN}"

  res = Net::HTTP.start(uri.hostname, uri.port, use_ssl: true) { |http| http.request(req) }
  [JSON.parse(res.body), res.each_header.to_h]
rescue => e
  puts "[에러] 멘션 불러오기 실패: #{e.message}"
  [[], {}]
end

def reply_to_mention(content, in_reply_to_id)
  uri = URI(POST_ENDPOINT)
  req = Net::HTTP::Post.new(uri)
  req['Authorization'] = "Bearer #{TOKEN}"
  req.set_form_data('status' => content, 'in_reply_to_id' => in_reply_to_id, 'visibility' => 'unlisted')

  Net::HTTP.start(uri.hostname, uri.port, use_ssl: true) { |http| http.request(req) }
rescue => e
  puts "[에러] 답글 전송 실패: #{e.message}"
end

REPLY = proc do |arg1, arg2|
  raw_content = arg1
  raw_id      = arg2

  puts "[REPLY ARGS] raw_content=#{raw_content.inspect} raw_id=#{raw_id.inspect}"

  content_str = nil
  id_str      = nil

  # 1) 첫 번째 인자가 status 객체(OpenStruct or Hash)인 경우
  if arg1.is_a?(OpenStruct) || arg1.is_a?(Hash)
    status = arg1
    toot_id =
      if status.respond_to?(:id)
        status.id
      else
        status['id']
      end

    content_str = arg2.to_s      # 두 번째 인자를 답글 내용으로
    id_str      = toot_id.to_s   # status.id를 in_reply_to_id로 사용

  else
    # 2) 일반적인 경우: (content, in_reply_to_id) 또는 (id, content)
    content_str = arg1.to_s
    id_str      = arg2.to_s

    # ⚡ 순서가 뒤집힌 경우 자동 교정:
    #   - id_str가 숫자가 아니고
    #   - content_str이 숫자처럼 생겼으면 -> 뒤집힌 걸로 간주
    if id_str !~ /^\d+$/ && content_str =~ /^\d+$/
      puts "[REPLY GUARD] 인자 순서가 뒤집혀 있어 교정합니다. (content ↔ id)"
      content_str, id_str = id_str, content_str
    end
  end

  # 최종적으로 in_reply_to_id가 숫자여야만 요청을 보냄
  unless id_str =~ /^\d+$/
    puts "[REPLY ERROR] in_reply_to_id가 숫자가 아닙니다: #{id_str.inspect}"
    return
  end

  uri = URI(POST_ENDPOINT)
  req = Net::HTTP::Post.new(uri)
  req['Authorization'] = "Bearer #{TOKEN}"
  req.set_form_data(
    'status'         => content_str,
    'in_reply_to_id' => id_str,
    'visibility'     => 'unlisted'
  )

  res = Net::HTTP.start(uri.hostname, uri.port, use_ssl: true) { |http| http.request(req) }

  puts "[REPLY] code=#{res.code} body=#{res.body[0,200].inspect} to=#{id_str}"
  if res.code.to_i >= 300
    puts "[에러] 답글 전송 실패 (HTTP #{res.code})"
  else
    puts "[REPLY OK] toot posted"
  end
rescue => e
  puts "[에러] 답글 전송 예외: #{e.class} - #{e.message}"
end


# ============================================
# Mentions 감시 루프 (Rate-limit 완전 대응)
# ============================================
last_checked_id = File.exist?(LAST_ID_FILE) ? File.read(LAST_ID_FILE).strip : nil
base_interval = 60
cooldown_on_429 = 300
loop_count = 0

puts "[MENTION] 감시 시작..."

loop do
  begin
    loop_count += 1
    delay = base_interval + rand(-10..10)
    puts "[루프 #{loop_count}] Mentions 확인 (지연 #{delay}s)"
    mentions, headers = fetch_mentions(last_checked_id)

    if headers['x-ratelimit-remaining'] && headers['x-ratelimit-remaining'].to_i < 1
      reset_after = headers['x-ratelimit-reset'] ? headers['x-ratelimit-reset'].to_i : cooldown_on_429
      puts "[경고] Rate limit 도달 → #{reset_after}초 대기"
      sleep(reset_after)
      next
    end

    mentions.sort_by! { |m| m['id'].to_i }
    mentions.each do |mention|
      next unless mention['type'] == 'mention'
      next unless mention['status']

      status = mention['status']
      sender = mention['account']['acct']
      content = status['content'].gsub(/<[^>]*>/, '').strip
      toot_id = status['id']

      puts "[MENTION] @#{sender}: #{content}"
      begin
        mention['status']  = OpenStruct.new(status)
        mention['account'] = OpenStruct.new(mention['account'])
        ProfessorParser.parse(REPLY, sheet_manager, mention)
      rescue => e
        puts "[에러] 명령어 실행 실패: #{e.message}"
      end

      last_checked_id = mention['id']
      File.write(LAST_ID_FILE, last_checked_id)
    end

  rescue => e
    if e.message.include?('429')
      puts "[경고] 429 Too Many Requests → 5분 대기"
      sleep(cooldown_on_429)
    else
      puts "[에러] Mentions 루프 오류: #{e.class} - #{e.message}"
      sleep(30)
    end
  end

  sleep(base_interval + rand(-10..10))
end
