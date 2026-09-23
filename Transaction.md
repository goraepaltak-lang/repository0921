⚙️ 핵심 비즈니스 로직 (Core Transactions)

본 프로젝트는 기존 수작업 위주의 항만 예인 업무를 자동화하기 위해 두 가지 핵심 트랜잭션을 처리합니다.


# 1. AI 기반 예인선 자동 배정 (AI Auto-Assignment)당직 순번 확인과 예인선 배정을 AI 알고리즘으로 자동화하며, 동시 다발적인 배정 요청 시 데이터 충돌을 방지(Concurrency Control)합니다.
    1-1. 관련 테이블: WORK_ORDER, VESSEL, RULE, DUTY_QUEUE, TUGBOAT, ASSIGNMENT
    1-2. 트랜잭션 흐름:정보 조회: 신규 작업 요청 시 선박(VESSEL)의 총톤수(G/T) 확인
    1-3. 규칙 적용: 매핑 규칙(RULE)에 따라 필요한 타겟 마력군과 척수 산출
    1-4. 동시성 제어: 당직 큐(DUTY_QUEUE) 조회 시 Lock(FOR UPDATE)을 걸어 순번 꼬임 방지
    1-5. 배정 처리: 예인선 배정 상세 내역(ASSIGNMENT)을 저장하고 예인선 및 작업 상태 업데이트


# 2. 영수증 자동 발급 (Automatic Invoice Generation)예인 작업 완료 시점에 복잡한 할증 로직(기상, 선박 특성, 시간대 등)을 자동으로 계산하여 청구서를 발행합니다.
    
    2-1. 관련 테이블: WORK_ORDER, VESSEL, INVOICE, SURCHARGE_DETAIL
    
    2-2. 트랜잭션 흐름:작업 완료 감지: 작업 지시(WORK_ORDER) 상태가 '완료'로 변경됨을 트리거로 트랜잭션 시작
    
    2-3. 할증 조건 판별: 위험물/초대형 선박 여부(VESSEL), 기상 악화 여부(WORK_ORDER) 등 파악
    
    2-4. 청구서 발행: 기본 요금과 할증 총액을 계산하여 마스터 청구서(INVOICE) 생성
    
    2-5. 상세 내역 기록: 적용된 모든 할증 내역을 분리하여 상세 테이블(SURCHARGE_DETAIL)에 기록   
