"""Approved message text and the three report types."""

TEST_NONE = "없음"


TEST_WEEKLY = "주간"


TEST_MONTHLY = "월간"


TEST_TYPES = (TEST_NONE, TEST_WEEKLY, TEST_MONTHLY)


TEMPLATE_TEXTS = {
    TEST_NONE: """[{academy_name}]
{student_name} 학생의 오늘 학습 현황을 안내드립니다.

■ 등원
{attendance}

■ 과제
수행률 {homework_percent}
미흡과제: {missing_homework}

■ 단어 테스트
{word_total_count}개 중 {word_wrong_count}개 오답

■ 오답 시험
{error_test_status}""",
    TEST_WEEKLY: """[{academy_name}]
{student_name} 학생의 오늘 학습 현황을 안내드립니다.

■ 등원
{attendance}

■ 과제
수행률 {homework_percent}
미흡과제: {missing_homework}

■ 단어 테스트
{word_total_count}개 중 {word_wrong_count}개 오답

■ 오답 시험
{error_test_status}

■ 주간 테스트
레벨: Lv{test_level}
점수: {test_score}점
평균: {test_average}점""",
    TEST_MONTHLY: """[{academy_name}]
{student_name} 학생의 오늘 학습 현황을 안내드립니다.

■ 등원
{attendance}

■ 과제
수행률 {homework_percent}
미흡과제: {missing_homework}

■ 단어 테스트
{word_total_count}개 중 {word_wrong_count}개 오답

■ 오답 시험
{error_test_status}

■ 월간 테스트
레벨: Lv{test_level}
점수: {test_score}점
평균: {test_average}점""",
}


COMMON_REQUIRED_VARIABLES = {
    "attendance": "등원상태",
    "homework_percent": "과제수행률",
    "word_total_count": "단어전체수",
    "word_wrong_count": "단어오답수",
    "error_test_status": "오답시험상태",
}
