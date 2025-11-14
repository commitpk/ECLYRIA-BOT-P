# ============================================
# /root/mastodon_bots/professor_bot/commands/attendance_command.rb
# ============================================
require_relative '../utils/professor_control'
require_relative '../utils/house_score_updater'
require 'date'

class AttendanceCommand
  include HouseScoreUpdater

  def initialize(sheet_manager, mastodon_client, sender, status)
    @sheet_manager = sheet_manager
    @mastodon_client = mastodon_client
    @sender = sender.gsub('@', '')
    @status = status
  end

  def execute
    puts "[Attendance] execute START for #{@sender}"

    # 1. 학생 정보 확인
    user = @sheet_manager.find_user(@sender)
    if user.nil?
      puts "[Attendance] user not found -> reply(입학 안내)"
      return professor_reply("아직 학적부에 없는 학생이군요. [입학/이름]으로 등록을 마쳐주세요.")
    end

    # 2. 출석 기능 상태 확인
    enabled = ProfessorControl.auto_push_enabled?(@sheet_manager, "아침출석자동툿")
    puts "[Attendance] 출석 기능 상태 = #{enabled}"
    unless enabled
      puts "[Attendance] 출석 기능 OFF -> reply(중단 안내)"
      return professor_reply("지금은 출석 기능이 잠시 중단된 상태예요. 나중에 다시 시도해보세요.")
    end

    today = Date.today.to_s
    current_time = Time.now
    puts "[Attendance] today=#{today}, user_attendance_date=#{user[:attendance_date].inspect}"

    # 3. 중복 출석 방지
    if user[:attendance_date] == today
      puts "[Attendance] 이미 오늘 출석함 -> reply(이미 출석 완료)"
      return professor_reply("오늘은 이미 출석을 완료했어요. 성실하군요, 훌륭합니다.")
    end

    # 4. 출석 가능 시간 확인 (22시 이전)
    if current_time.hour >= 22
      puts "[Attendance] 22시 이후 -> reply(마감 안내)"
      return professor_reply("출석 마감 시간(22:00)이 지나버렸군요. 내일은 조금 더 일찍 오도록 해요.")
    end

    # 5. 출석 처리
    puts "[Attendance] 출석 처리 진행 (갈레온/기숙사점수 갱신)"
    @sheet_manager.increment_user_value(@sender, "갈레온", 2)
    @sheet_manager.increment_user_value(@sender, "기숙사점수", 1)
    @sheet_manager.set_user_value(@sender, "출석날짜", today)

    # 6. 기숙사 점수 반영
    update_house_scores(@sheet_manager)

    # 7. 교수님식 출석 멘트
    user_name = user[:name] || @sender
    message = "좋아요, #{user_name} 학생. 오늘도 성실히 출석했군요.\n(보상: 2갈레온, 기숙사 점수 +1)"
    puts "[Attendance] 정상 출석 -> reply(출석 완료 멘션)"
    professor_reply(message)

  rescue => e
    puts "[에러] AttendanceCommand 처리 중 예외 발생: #{e.message}"
    puts e.backtrace
    professor_reply("음… 잠시 오류가 생긴 것 같아요. 잠시 후 다시 시도해보세요.")
  end


  private

  private

  def professor_reply(message)
    message = message.to_s.empty? ? "출석이 확인되었습니다." : message.dup
  
    # status가 OpenStruct인지, Hash인지에 따라 id 꺼내기
    status_id =
      if @status.respond_to?(:id)
        @status.id
      else
        (@status['id'] rescue @status[:id])
      end
    
    puts "[Attendance] professor_reply(to_status_id=#{status_id}, msg=#{message.inspect})"
    @mastodon_client.reply(message, status_id)
  end
end
