# ============================================
# sheet_manager.rb (교수봇용 안정화 버전 - 안전 A1 유틸 포함)
# ============================================
require 'google/apis/sheets_v4'

class SheetManager
  attr_reader :service, :sheet_id

  # 시트 탭 이름(필요하면 여기만 바꾸면 됩니다)
  USERS_SHEET = '사용자'.freeze
  PROFESSOR_SHEET = '교수'.freeze # <--- 상수를 클래스 내부로 이동

  def initialize(service, sheet_id)
    @service = service
    @sheet_id = sheet_id
  end

  def read(sheet_name, a1 = 'A:Z')
    ensure_separate_args!(sheet_name, a1)
    read_range(a1_range(sheet_name, a1))
  end

  def write(sheet_name, a1, values)
    ensure_separate_args!(sheet_name, a1)
    write_range(a1_range(sheet_name, a1), values)
  end

  def append(sheet_name, row)
    ensure_separate_args!(sheet_name, 'A:Z')
    append_log(sheet_name, row)
  end

  # ---------- A1 유틸 ----------
  # 시트 이름에 한글/공백/특수문자/작은따옴표가 있어도 안전하게 A1 범위를 만듭니다.
  def a1_range(sheet_name, a1 = 'A:Z')
    sh = sheet_name.to_s
    if sh.include?('!')
      base, rng_from_name = sh.split('!', 2)
      # a1이 명시되었고 기본값이 아니라면 a1 우선, 아니면 시트명에 들어온 범위 사용
      rng = (a1 && a1.strip != '' && a1 != 'A:Z') ? a1 : rng_from_name
      escaped = base.gsub("'", "''")  # ' → ''
      "'#{escaped}'!#{rng}"
    else
      escaped = sh.gsub("'", "''")
      "'#{escaped}'!#{a1}"
    end
  end

  # 0-based column index -> A1 column letters (0->A, 25->Z, 26->AA ...)
  def col_idx_to_a1(idx)
    s = ''
    n = idx
    while n >= 0
      s = (65 + (n % 26)).chr + s
      n = (n / 26) - 1
    end
    s
  end

  # ---------- 공통 I/O ----------
  # 시트의 특정 범위 읽기
  def read_range(range)
    response = @service.get_spreadsheet_values(@sheet_id, range)
    response.values || []
  rescue => e
    puts "[시트 읽기 오류] #{e.message}"
    []
  end

  # 시트의 특정 범위 쓰기
  def write_range(range, values)
    value_range = Google::Apis::SheetsV4::ValueRange.new(values: values)
    @service.update_spreadsheet_value(
      @sheet_id,
      range,
      value_range,
      value_input_option: 'USER_ENTERED'
    )
  rescue => e
    puts "[시트 쓰기 오류] #{e.message}"
  end

  # 로그 남기기 (예: 출석, 과제 기록)
  def append_log(sheet_name, row)
    range = a1_range(sheet_name, 'A:Z')
    value_range = Google::Apis::SheetsV4::ValueRange.new(values: [row])
    @service.append_spreadsheet_value(
      @sheet_id,
      range,
      value_range,
      value_input_option: 'USER_ENTERED'
    )
  rescue => e
    puts "[시트 로그 추가 오류] #{e.message}"
  end

  def append_values(range, values)
    vr = Google::Apis::SheetsV4::ValueRange.new(values: values)
    @service.append_spreadsheet_value(
      @sheet_id,
      range,
      vr,
      value_input_option: 'USER_ENTERED'
    )
  rescue => e
    puts "[append_values 오류] #{e.class} - #{e.message}"
  end

  # ============================================
  # 학적부 관리 기능
  # ============================================

  # 특정 유저 찾기
  def find_user(username)
    data = read_range(a1_range(USERS_SHEET, 'A:Z'))
    return nil if data.empty?

    header = data[0] || []
    return nil if data.size < 2

    username_col     = header.index('아이디') || 0
    name_col         = header.index('이름')   || 1
    galleon_col      = header.index('갈레온')
    house_score_col  = header.index('개별 기숙사 점수')
    attend_col       = header.index('출석날짜')

    row = data.find.with_index { |r, i| i > 0 && r[username_col].to_s.strip == username.strip }
    return nil unless row

    {
      id:              row[username_col],
      name:            row[name_col],
      galleon:         galleon_col     ? row[galleon_col].to_i   : 0,
      house_score:     house_score_col ? row[house_score_col].to_i : 0,
      attendance_date: attend_col      ? (row[attend_col].to_s)   : ''
    }
  rescue => e
    puts "[find_user 오류] #{e.message}"
    nil
  end

  # 유저의 특정 열 값을 증가시킴
  def increment_user_value(username, column_name, value)
    data = read_range(a1_range(USERS_SHEET, 'A:Z'))
    return if data.empty?
    header = data[0] || []

    target_col = header.index(column_name)
    return if target_col.nil?

    username_col = header.index('아이디') || 0

    data.each_with_index do |row, i|
      next if i.zero?
      next unless row[username_col].to_s.strip == username.strip

      current = (row[target_col] || 0).to_i
      new_val = current + value

      col_letter = col_idx_to_a1(target_col)
      cell_range = a1_range(USERS_SHEET, "#{col_letter}#{i + 1}")

      write_range(cell_range, [[new_val]])
      puts "[시트 업데이트] #{username}의 #{column_name} → #{new_val}"
      return
    end
  rescue => e
    puts "[increment_user_value 오류] #{e.message}"
  end

  # 유저의 특정 열 값을 설정
  def set_user_value(username, column_name, new_value)
    data = read_range(a1_range(USERS_SHEET, 'A:Z'))
    return if data.empty?
    header = data[0] || []

    target_col = header.index(column_name)
    return if target_col.nil?

    username_col = header.index('아이디') || 0

    data.each_with_index do |row, i|
      next if i.zero?
      next unless row[username_col].to_s.strip == username.strip

      col_letter = col_idx_to_a1(target_col)
      cell_range = a1_range(USERS_SHEET, "#{col_letter}#{i + 1}")

      write_range(cell_range, [[new_value]])
      puts "[시트 설정] #{username}의 #{column_name} = #{new_value}"
      return
    end
  rescue => e
    puts "[set_user_value 오류] #{e.message}"
  end

  # ============================================
  # 자동 푸시 여부 확인 기능
  # ============================================
  def auto_push_enabled?(key: '아침출석자동툿', key_col: '설정', val_col: '값')
    range = a1_range(PROFESSOR_SHEET, 'A1:Z2')
    data  = read_range(range)

    puts "[DEBUG/auto_push] 읽은 범위: #{range.inspect}"
    puts "[DEBUG/auto_push] 읽은 데이터: #{data.inspect}"

    return false if data.empty? || data[0].nil?

    header = data[0]
    values = data[1] || []

    # 헤더 정규화 (공백·전각문자 제거)
    normalized_key = key.to_s.strip.unicode_normalize(:nfkc)
    header_index = header.index { |h| h.to_s.strip.unicode_normalize(:nfkc) == normalized_key }

    puts "[DEBUG/auto_push] 찾은 키 인덱스: #{header_index.inspect}"

    return false if header_index.nil?

    val = values[header_index]
    puts "[DEBUG/auto_push] 읽은 값: #{val.inspect} (클래스=#{val.class})"

    # 체크박스가 boolean으로 오거나 문자열로 올 수 있음
    if val == true || val.to_s.strip.upcase == 'TRUE' || %w[ON YES ✅ ☑ 1].include?(val.to_s.strip.upcase)
      puts "[DEBUG/auto_push] 최종 판정: true"
      true
    else
      puts "[DEBUG/auto_push] 최종 판정: false"
      false
    end
  rescue => e
    puts "[auto_push_enabled? 오류] #{e.message}"
    false
  end


  # ============================================
  # PRIVATE 유틸리티
  # ============================================
  private # <--- 내부에서만 사용하는 메서드는 private으로 지정

  # 누락되었던 필수 메서드 정의 추가
  def ensure_separate_args!(sheet_name, a1)
    unless sheet_name.is_a?(String) && !sheet_name.strip.empty?
      raise ArgumentError, "시트 이름이 유효하지 않습니다."
    end
    unless a1.is_a?(String) && !a1.strip.empty?
      raise ArgumentError, "A1 범위가 유효하지 않습니다."
    end
  end
end