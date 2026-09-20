import React, { useState, useMemo, useEffect, useRef, useId } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";
import "@coreui/coreui/dist/css/coreui.min.css";
import { CHeader, CHeaderBrand, CHeaderNav, CHeaderToggler, CContainer, CSidebar, CSidebarNav, CNavItem, CNavGroup } from "@coreui/react";
import CIcon from "@coreui/icons-react";
import { cilX, cilWarning, cilLightbulb, cilSpeech, cilCog, cilLink } from "@coreui/icons";

/* ============================================================
   Career OS MVP — v1.2 프로토타입
   - 더미 데이터 / 상태관리 useState / 데스크톱 우선
   - 핵심 5단계 + 심화 4단계 분석 플로우
   - 수치 토큰 {{metric:id|format}} 렌더링
   - 승인 상태(approvalStatus) / isStale 표시
   - 역량·스킬 탭 (활용 범위 + 경험 근거 연결)
   - 파일 가져오기: 비정형 텍스트/문서 → AI 추출 → 사용자 확인·수정 후 반영
   ============================================================ */

/* ---------- 디자인 토큰 — CoreUI Free React Admin Template 스타일 ---------- */
const C = {
  // 값은 src/styles.css 의 :root 에 있다. 여기서는 참조만 한다.
  // 인라인 style 에 var() 문자열을 그대로 넣어도 브라우저가 해석한다.
  // 표면
  bg: "var(--c-bg)", panel: "var(--c-panel)", line: "var(--c-line)", lineSoft: "var(--c-line-soft)", accent: "var(--c-accent)",
  // 텍스트
  text: "var(--c-text)", sub: "var(--c-sub)", faint: "var(--c-faint)",
  // 의미색 — 채움용
  blue: "var(--c-blue)", blueBg: "var(--c-blue-bg)", primary: "var(--c-primary)", primaryBg: "var(--c-primary-bg)", green: "var(--c-green)", greenBg: "var(--c-green-bg)", orange: "var(--c-orange)", orangeBg: "var(--c-orange-bg)", ai: "var(--c-ai)", aiBg: "var(--c-ai-bg)", red: "var(--c-red)", redBg: "var(--c-red-bg)",
  // 의미색 — 텍스트용 (WCAG AA)
  faintText: "var(--c-faint-text)", blueText: "var(--c-blue-text)", greenText: "var(--c-green-text)", orangeText: "var(--c-orange-text)", redText: "var(--c-red-text)",
  // 사이드바 (다크)
  sidebarBg: "var(--c-sidebar-bg)", sidebarText: "var(--c-sidebar-text)", sidebarTextActive: "var(--c-sidebar-text-active)", sidebarHover: "var(--c-sidebar-hover)", sidebarLine: "var(--c-sidebar-line)",
};
const font = "'Spoqa Han Sans Neo','Spoqa Han Sans JP',system-ui,-apple-system,'Segoe UI',Roboto,'Helvetica Neue','Apple SD Gothic Neo','Noto Sans KR',Arial,sans-serif";

// Spoqa Han Sans Neo 웹폰트 로드 — index.html이 없는 실행 환경(아티팩트 미리보기 등)에서도
// 런타임에 스타일시트를 주입해 폰트가 적용되도록 함. 이미 <head>에 있으면 중복 삽입하지 않음.
function useSpoqaHanSansFont() {
  useEffect(() => {
    try {
      const href = "https://spoqa.github.io/spoqa-han-sans/css/SpoqaHanSansNeo.css";
      if (document.querySelector(`link[href="${href}"]`)) return;
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = href;
      document.head.appendChild(link);
    } catch { /* 폰트 로드 실패해도 시스템 폰트로 정상 동작 */ }
  }, []);
}

/* ---------- 예시 데이터 (일반적인 아르바이트·인턴 경험 기준) ---------- */
const seedMetrics = [
  { id: "m_001", experienceId: "e_1", metricType: "revenue", metricName: "매출", changeValue: 29, unit: "%", comparisonBasis: "전년 동기간", evidenceSource: "정산 리포트", certainty: "verified", isPublic: true },
  { id: "m_002", experienceId: "e_1", metricType: "orders", metricName: "주문 수", changeValue: 36, unit: "%", comparisonBasis: "전년 동기간", evidenceSource: "정산 리포트", certainty: "verified", isPublic: true },
  { id: "m_003", experienceId: "e_2", metricType: "time", metricName: "재고 확인 소요 시간", beforeValue: 180, afterValue: 5, unit: "분", certainty: "verified", isPublic: true },
  { id: "m_004", experienceId: "e_2", metricType: "operation", metricName: "출고 오류·CS", afterValue: 0, unit: "건", certainty: "verified", isPublic: true },
  { id: "m_005", experienceId: "e_3", metricType: "revenue", metricName: "채널 매출", changeValue: 42, unit: "%", comparisonBasis: "개편 전 3개월", certainty: "memory_based", isPublic: true },
];

const seedExperiences = [
  {
    id: "e_1", title: "웹사이트 운영 프로모션 기획", organization: "온라인 쇼핑몰 (인턴)", experienceType: "internship", primaryCategory: "온라인 쇼핑몰 인턴",
    startDate: "2024-06", endDate: "2024-08", role: "이커머스 운영 인턴",
    rawNote: "인턴 근무 중 시즌 프로모션을 진행했다. 기존 할인만 하는 것보다 사은품을 주는 게 좋을 것 같았다. 과거 데이터를 분석해서 제품을 골랐고 매출이 올랐다.",
    context: "시즌 최대 행사를 앞두고 이커머스팀 4명이 프로모션을 준비하는 상황",
    assignedTask: "시즌 프로모션 운영", discoveredProblem: "단순 할인만으로는 객단가와 연관 구매를 높이기 어려움",
    goal: "매출·주문 수 증가, 객단가 유지 또는 상승",
    personalContribution: "과거 3년 판매·장바구니 데이터를 분석하고, 인기 제품 3종과 조건부 증정 구조를 제안한 뒤 세팅과 성과 분석까지 담당",
    contributionLevel: "proposed_and_executed",
    contributionEvidence: "과거 3년 판매 데이터를 직접 분석하고, 사은품 제품과 구매 조건을 제안한 뒤, 프로모션 세팅과 결과 리포트까지 담당했다.",
    difficulty: "행사 2주 전 확정이라 분석 시간이 부족했고, 증정 재고 부담에 대한 매니저 우려를 설득해야 했다",
    learning: "할인 폭보다 구매 조건 설계가 객단가를 움직인다는 것을 확인 → 이후 모든 프로모션에서 장바구니 데이터를 먼저 확인하는 습관",
    jobRelevance: "MD·이커머스 직무의 상품 구성·딜 설계 역량과 직결",
    coreMessage: "데이터를 단순히 보고하는 데 그치지 않고, 고객의 구매 흐름을 설계해 매출 성과로 연결했다.",
    oneLineSummary: "3년간의 판매 데이터를 분석해 조건부 증정 프로모션을 설계하고, 전년 동기간 대비 매출 29%·주문 36% 증가를 달성",
    status: "complete", depthDone: true, usageCount: 5, updatedAt: "2026-07-18",
    competencies: ["데이터 분석", "프로모션 기획", "고객 분석", "문제 해결", "실행력"],
    tags: ["개인+팀", "성공", "정량 성과"],
    actions: [
      { id: "a1", actionType: "analysis", description: "과거 3년 판매량, 장바구니, 상품별 성과 분석", isDirectAction: true },
      { id: "a2", actionType: "judgment", description: "구매 가능성이 높은 인기 제품 3종 선정", isDirectAction: true },
      { id: "a3", actionType: "execution", description: "조건부 증정 구조와 연관 상품 배치 설계", isDirectAction: true },
      { id: "a4", actionType: "collaboration", description: "디자인·마케팅팀에 프로모션 콘텐츠 요청", isDirectAction: true },
    ],
    completion: { 배경: "충분", 문제: "충분", 행동: "충분", 기여도: "충분", 성과: "충분", 목표: "충분", 어려움: "충분", "배운 점": "충분", "직무 연결": "보완 필요" },
  },
  {
    id: "e_2", title: "재고 관리 업무 자동화", organization: "온라인 쇼핑몰 (인턴)", experienceType: "internship", primaryCategory: "온라인 쇼핑몰 인턴",
    startDate: "2024-07", endDate: "2024-09", role: "이커머스 운영 인턴",
    rawNote: "재고 시스템과 판매 플랫폼이 연동이 안 돼서 품절 상품이 계속 노출됐다. 엑셀 VBA를 배워서 자동화 파일을 만들었다.",
    context: "다수의 상품을 매주 수작업으로 대조하던 상황",
    assignedTask: "주간 재고 확인", discoveredProblem: "시스템 간 미연동으로 품절 상품이 노출되어 주문 취소·CS 반복",
    goal: "품절 상품을 빠르게 구분하고 출고 오류와 주문 취소를 줄이는 것",
    personalContribution: "Excel 함수와 VBA를 학습해 자동화 파일을 제작하고, 주간 점검 프로세스와 매뉴얼을 만들었다.",
    contributionLevel: "led", contributionEvidence: "문제 발견부터 자동화 파일 제작, 매뉴얼 배포까지 단독 수행",
    difficulty: "", learning: "", jobRelevance: "",
    coreMessage: "반복 업무를 문제로 정의하고 스스로 도구를 학습해 구조적으로 해결했다.",
    oneLineSummary: "Excel VBA 자동화로 재고 확인 시간을 180분→5분으로 단축하고 출고 오류 0건 달성",
    status: "complete", depthDone: false, usageCount: 3, updatedAt: "2026-07-15",
    competencies: ["문제 해결", "업무 자동화", "데이터 관리", "주도성", "운영 개선"],
    tags: ["개인", "성공", "정량 성과"],
    actions: [
      { id: "a5", actionType: "analysis", description: "재고 시스템 간 데이터를 비교해 품절 발생 원인 분석", isDirectAction: true },
      { id: "a6", actionType: "judgment", description: "다수 상품 수작업 비교는 지속 불가능하다고 판단", isDirectAction: true },
      { id: "a7", actionType: "execution", description: "Excel 함수·VBA로 품절 상품 자동 구분 파일 제작", isDirectAction: true },
      { id: "a8", actionType: "improvement", description: "주간 점검 프로세스와 사용 매뉴얼 배포", isDirectAction: true },
    ],
    completion: { 배경: "충분", 문제: "충분", 행동: "충분", 기여도: "충분", 성과: "충분", 목표: "미입력", 어려움: "미입력", "배운 점": "미입력", "직무 연결": "미입력" },
  },
  {
    id: "e_3", title: "온라인 채널 콘텐츠 개편", organization: "온라인 쇼핑몰 (인턴)", experienceType: "internship", primaryCategory: "온라인 쇼핑몰 인턴",
    startDate: "2024-08", endDate: "2024-10", role: "이커머스 운영 인턴",
    rawNote: "방치돼 있던 온라인 스토어를 다시 살렸다. 상세 콘텐츠를 만들고 리스팅을 정리했다.",
    context: "", assignedTask: "", discoveredProblem: "리스팅 이미지·키워드가 방치되어 노출·전환 모두 하락",
    goal: "", personalContribution: "상세 콘텐츠 제작과 리스팅 최적화를 직접 제안하고 실행",
    contributionLevel: "proposed_and_executed", contributionEvidence: "",
    coreMessage: "", oneLineSummary: "방치된 온라인 채널을 재정비해 매출 회복",
    status: "needs_revision", depthDone: false, usageCount: 1, updatedAt: "2026-07-10",
    competencies: ["채널 운영", "콘텐츠 기획", "주도성"],
    tags: ["개인", "성공", "정량 성과"],
    actions: [
      { id: "a9", actionType: "analysis", description: "경쟁 리스팅 대비 이미지·키워드 격차 분석", isDirectAction: true },
      { id: "a10", actionType: "execution", description: "상세 콘텐츠 제작 및 리스팅 재정비", isDirectAction: true },
    ],
    completion: { 배경: "보완 필요", 문제: "충분", 행동: "보완 필요", 기여도: "보완 필요", 성과: "보완 필요", 목표: "미입력", 어려움: "미입력", "배운 점": "미입력", "직무 연결": "미입력" },
  },
  {
    id: "e_4", title: "동아리 콘텐츠팀 운영", organization: "대학 홍보 동아리", experienceType: "club", primaryCategory: "동아리 활동",
    startDate: "2023-03", endDate: "2023-12", role: "콘텐츠팀장",
    rawNote: "콘텐츠팀장으로 6명 팀을 운영했다. 업로드 일정이 계속 밀리는 문제가 있었다.",
    context: "", assignedTask: "", discoveredProblem: "", goal: "", personalContribution: "",
    contributionLevel: "", contributionEvidence: "", coreMessage: "", oneLineSummary: "",
    status: "draft", depthDone: false, usageCount: 0, updatedAt: "2026-07-05",
    competencies: ["리더십"], tags: ["팀"], actions: [],
    completion: { 배경: "미입력", 문제: "미입력", 행동: "미입력", 기여도: "미입력", 성과: "미입력", 목표: "미입력", 어려움: "미입력", "배운 점": "미입력", "직무 연결": "미입력" },
  },
];

const seedOutputs = [
  { id: "o_1", experienceId: "e_1", outputType: "resume", style: "result_focused",
    content: "3년간 판매 데이터를 분석해 조건부 증정 프로모션을 기획하고, 전년 동기간 대비 매출 {{metric:m_001|exact}}·주문 수 {{metric:m_002|exact}} 증가 달성",
    referencedMetricIds: ["m_001", "m_002"], version: 2, isAiGenerated: true, approvalStatus: "approved", isStale: false },
  { id: "o_2", experienceId: "e_1", outputType: "resume", style: "role_focused",
    content: "온라인 채널 운영과 판매 데이터 분석, 프로모션 기획 담당",
    referencedMetricIds: [], version: 1, isAiGenerated: true, approvalStatus: "ai_draft", isStale: false },
  { id: "o_3", experienceId: "e_1", outputType: "interview",
    content: "데이터 분석으로 매출 성과를 낸 경험을 말씀드리겠습니다. 인턴 당시 시즌 프로모션에서 단순 할인의 한계를 발견하고, 3년치 판매·장바구니 데이터를 분석해 조건부 증정 구조를 제안·실행했습니다. 그 결과 매출 {{metric:m_001|rounded}}, 주문 수 {{metric:m_002|rounded}} 증가를 달성했습니다.",
    referencedMetricIds: ["m_001", "m_002"], version: 1, isAiGenerated: true, approvalStatus: "user_editing", isStale: false },
  { id: "o_4", experienceId: "e_2", outputType: "resume", style: "result_focused",
    content: "Excel VBA 기반 재고 자동화로 확인 시간 {{metric:m_003|exact}} 단축, 출고 오류 {{metric:m_004|exact}} 달성",
    referencedMetricIds: ["m_003", "m_004"], version: 1, isAiGenerated: true, approvalStatus: "approved", isStale: false },
  { id: "o_5", experienceId: "e_3", outputType: "resume", style: "result_focused",
    content: "방치된 온라인 채널을 재정비해 개편 전 대비 매출 {{metric:m_005|exact}} 회복",
    referencedMetricIds: ["m_005"], version: 1, isAiGenerated: true, approvalStatus: "ai_draft", isStale: true },
];

/* ---------- 스킬 · 자격증 예시 ---------- */
const seedSkills = [
  {
    id: "s_1", name: "쇼핑몰 관리 플랫폼", category: "tool",
    summary: "리스팅·프로모션·재고 연동까지 운영 전반",
    scopeItems: [
      { id: "sc1", text: "상품 리스팅 등록·수정, 컬렉션 구성", evidenceExpId: "e_3" },
      { id: "sc2", text: "조건부 증정 프로모션 세팅", evidenceExpId: "e_1" },
      { id: "sc3", text: "판매·장바구니 리포트 추출 및 분석", evidenceExpId: "e_1" },
      { id: "sc4", text: "재고 데이터 추출 후 시스템 간 대조", evidenceExpId: "e_2" },
      { id: "sc5", text: "테마·페이지 커스터마이징", evidenceExpId: null },
    ],
  },
  {
    id: "s_2", name: "Excel / VBA", category: "tool",
    summary: "함수·피벗은 능숙, VBA는 실무 자동화 1건 완수 수준",
    scopeItems: [
      { id: "sc6", text: "VLOOKUP·INDEX/MATCH·조건부 서식 등 실무 함수", evidenceExpId: "e_2" },
      { id: "sc7", text: "피벗 테이블 기반 판매 데이터 집계", evidenceExpId: "e_1" },
      { id: "sc8", text: "VBA 매크로로 재고 대조 자동화", evidenceExpId: "e_2" },
      { id: "sc9", text: "파워쿼리·대시보드 구축", evidenceExpId: null },
    ],
  },
  {
    id: "s_3", name: "온라인 채널 운영", category: "tool",
    summary: "리스팅 최적화·상세 콘텐츠 중심",
    scopeItems: [
      { id: "sc10", text: "상세 콘텐츠 기획·제작", evidenceExpId: "e_3" },
      { id: "sc11", text: "리스팅 키워드·이미지 최적화", evidenceExpId: "e_3" },
      { id: "sc12", text: "광고 운영", evidenceExpId: null },
    ],
  },
  {
    id: "s_4", name: "데이터 분석", category: "skill",
    summary: "판매·장바구니 데이터를 상품 구성 의사결정으로 연결",
    scopeItems: [
      { id: "sc13", text: "3년치 판매 데이터에서 상품 선정 근거 도출", evidenceExpId: "e_1" },
      { id: "sc14", text: "상품 단위 성과 비교로 문제 원인 분석", evidenceExpId: "e_2" },
    ],
  },
];
const seedCerts = [
  { id: "c_1", name: "컴퓨터활용능력 1급", issuer: "대한상공회의소", date: "2023-05", note: "" },
  { id: "c_2", name: "TOEIC 900", issuer: "ETS", date: "2025-11", note: "유효기간 2027-11" },
  { id: "c_3", name: "OPIc IH", issuer: "ACTFL", date: "", note: "응시 예정", planned: true },
];

const seedMasterEssays = [
  { id: "mq_1", question: "자기소개를 해주세요.", characterLimit: 800, draft: "", status: "not_started", chatHistory: [] },
  { id: "mq_2", question: "지원 동기는 무엇인가요?", characterLimit: 700, draft: "", status: "not_started", chatHistory: [] },
  { id: "mq_3", question: "입사 후 포부를 말씀해주세요.", characterLimit: 600, draft: "", status: "not_started", chatHistory: [] },
];
const seedMasterInterviews = [
  { id: "miq_1", question: "자신의 장단점은 무엇인가요?", category: "인성", selectedExperienceId: null, practiceCount: 0, confidence: null, followUps: [], draft: "", chatHistory: [] },
  { id: "miq_2", question: "팀워크를 발휘했던 경험을 말씀해주세요.", category: "협업", selectedExperienceId: null, practiceCount: 0, confidence: null, followUps: [], draft: "", chatHistory: [] },
  { id: "miq_3", question: "실패하거나 좌절했던 경험은 무엇인가요?", category: "실패", selectedExperienceId: null, practiceCount: 0, confidence: null, followUps: [], draft: "", chatHistory: [] },
];

const seedQuestionBlocks = [
  { id: "qb_1", label: "가장 큰 성과", expIds: ["e_1"] },
  { id: "qb_2", label: "주도적으로 개선한 경험", expIds: ["e_1", "e_2"] },
  { id: "qb_3", label: "협업 경험", expIds: ["e_1"] },
  { id: "qb_4", label: "리더십 경험", expIds: ["e_4"] },
  { id: "qb_5", label: "어려움을 극복한 경험", expIds: ["e_1"] },
  { id: "qb_6", label: "문제 해결 경험", expIds: ["e_2"] },
];
// 신규 사용자 기본값 — 예시 경험을 참조하지 않는 빈 질문 블록
const emptyQuestionBlocks = [
  { id: "qb_1", label: "가장 큰 성과", expIds: [] },
  { id: "qb_2", label: "주도적으로 개선한 경험", expIds: [] },
  { id: "qb_3", label: "협업 경험", expIds: [] },
  { id: "qb_4", label: "리더십 경험", expIds: [] },
  { id: "qb_5", label: "어려움을 극복한 경험", expIds: [] },
  { id: "qb_6", label: "문제 해결 경험", expIds: [] },
];

const seedApplications = [
  { id: "ap_1", company: "A 리테일 기업", position: "MD (상품기획)", deadline: "2026-08-03", status: "writing", priority: "high",
    essayProgress: 60, interviewProgress: 30,
    requirements: [
      { id: "r1", requirement: "매출 데이터 분석", category: "required_competency", importance: 5, matchedExp: "e_1", matchReason: "판매 분석을 매출 성과로 연결", gap: "" },
      { id: "r2", requirement: "운영 개선 경험", category: "experience", importance: 4, matchedExp: "e_2", matchReason: "재고 관리 자동화 경험", gap: "" },
      { id: "r3", requirement: "채널 운영 이해", category: "preferred_competency", importance: 3, matchedExp: "e_3", matchReason: "채널 콘텐츠·리스팅 개선 경험", gap: "경험 분석 보완 필요" },
      { id: "r4", requirement: "협상·소싱 경험", category: "required_competency", importance: 4, matchedExp: null, matchReason: "", gap: "매칭 가능한 경험 없음" },
    ],
    essays: [
      { id: "q1", question: "지원 동기와 입사 후 포부를 기술하시오.", characterLimit: 700, status: "drafting", selectedExperienceIds: ["e_1"], isLocked: false, chatHistory: [] },
      { id: "q2", question: "가장 큰 성과를 낸 경험을 기술하시오.", characterLimit: 1000, status: "complete", selectedExperienceIds: ["e_1"], isLocked: true, chatHistory: [] },
    ],
    interviews: [
      { id: "iq1", question: "본인이 데이터로 성과를 낸 경험은?", category: "achievement", selectedExperienceId: "e_1", practiceCount: 4, confidence: 4,
        followUps: ["왜 그 방법을 선택했나요?", "본인이 직접 한 부분은 무엇인가요?", "성과가 본인의 행동 때문이라는 근거는?"] },
      { id: "iq2", question: "실패하거나 아쉬웠던 경험은?", category: "failure", selectedExperienceId: null, practiceCount: 0, confidence: null, followUps: [] },
    ],
  },
  { id: "ap_2", company: "B 유통 기업", position: "상품기획 MD", deadline: "2026-08-20", status: "analyzing", priority: "medium",
    essayProgress: 0, interviewProgress: 0, requirements: [], essays: [], interviews: [] },
];

/* ---------- 상수 ---------- */
const CORE_STEPS = ["배경", "문제", "행동", "기여도", "성과"];
const DEPTH_STEPS = ["목표", "어려움", "배운 점", "직무 연결"];
const STEP_QUESTIONS = {
  배경: ["언제, 어디에서 한 경험인가?", "팀이나 조직의 목표는 무엇이었나?", "당시 맡은 공식 역할은 무엇이었나?", "경험이 시작되기 전 상황은 어땠나?"],
  문제: ["당시 해결해야 했던 문제는 무엇이었나?", "기존 방식에는 어떤 한계가 있었나?", "어떤 데이터나 현상을 보고 문제라고 판단했나?"],
  행동: ["가장 먼저 한 행동은 무엇인가?", "왜 그 방법을 선택했나?", "본인이 직접 결정한 것은 무엇인가?", "다른 사람에게 요청하거나 설득한 것은 무엇인가?"],
  기여도: ["팀 전체가 한 일과 본인이 한 일을 구분하면?", "내가 없었다면 결과가 어떻게 달라졌을까?", "제안만 했는가, 실행까지 담당했는가?"],
  성과: ["결과가 이전보다 어떻게 달라졌나?", "수치로 표현할 수 있는가?", "비교 기준은 무엇인가?", "성과를 증명할 자료가 있는가?"],
  목표: ["달성해야 했던 목표는 무엇이었나?", "성공 여부를 어떤 기준으로 판단했나?", "시간·비용·인력 제약은 있었나?"],
  어려움: ["가장 어려웠던 점은 무엇인가?", "팀원 또는 상사와 의견 차이가 있었나?", "시간·인력·예산·정보 중 무엇이 부족했나?"],
  "배운 점": ["이전에는 어떻게 생각했나?", "경험을 통해 무엇을 새롭게 알게 됐나?", "다시 한다면 무엇을 바꾸겠나?"],
  "직무 연결": ["이 경험은 어떤 직무와 연결되는가?", "어떤 역량을 보여주는가?", "이 경험의 핵심 메시지는 무엇인가?"],
};
const STATUS_LABEL = { draft: "초기 메모", analyzing: "분석 중", needs_revision: "보완 필요", complete: "분석 완료" };
const STATUS_COLOR = { draft: [C.sub, C.lineSoft], analyzing: [C.blueText, C.blueBg], needs_revision: [C.orangeText, C.orangeBg], complete: [C.greenText, C.greenBg] };
const CONTRIB_LABEL = { participated: "참여", responsible: "담당", led: "주도", proposed_and_executed: "제안 후 실행", full_ownership: "전체 책임" };
const ACTION_LABEL = { goal: "목표", analysis: "분석", judgment: "판단", execution: "실행", collaboration: "협업", improvement: "개선" };
// 행동 카드 전용 색상 (앱 전체는 무채색 기조지만, 유형 구분이 중요한 이 영역만 예외적으로 색을 씀)
const ACTION_COLOR = {
  goal: ["#5E6997", "#EAECF5"],
  analysis: ["#2F6FA8", "#E7F0F7"],
  judgment: ["#7A5AA8", "#EFEAF6"],
  execution: ["#3D7659", "#E7F1EA"],
  collaboration: ["#AE4873", "#FBEAF0"],
  improvement: ["#8F6315", "#FBF1DF"],
};
const APPROVAL = {
  ai_draft: { label: "AI 초안 · 미승인", color: C.ai, bg: C.aiBg },
  user_editing: { label: "수정 중", color: C.blueText, bg: C.blueBg },
  approved: { label: "승인됨", color: C.greenText, bg: C.greenBg },
  rejected: { label: "폐기", color: C.faintText, bg: C.lineSoft },
};
const CERTAINTY = { verified: ["자료로 확인됨", C.greenText, C.greenBg], memory_based: ["기억에 기반함", C.orangeText, C.orangeBg], estimated: ["추정치", C.orangeText, C.orangeBg], needs_verification: ["추가 확인 필요", C.redText, C.redBg] };

/* ---------- 카테고리 선택 (드롭다운 + 새 카테고리 추가) ---------- */
function CategorySelect({ value, options, onChange, onAddOption, placeholder, style }) {
  const [adding, setAdding] = useState(false);
  const [draft, setDraft] = useState("");

  const commit = () => {
    const t = draft.trim();
    if (t) { onAddOption(t); onChange(t); }
    setAdding(false); setDraft("");
  };

  if (adding) {
    return (
      <div style={{ display: "flex", gap: 6, alignItems: "center" }} className="wrap-sm">
        <Input autoFocus placeholder="새 카테고리명" value={draft} onChange={e => setDraft(e.target.value)}
          onKeyDown={e => e.key === "Enter" && commit()} style={{ width: 140, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)", ...style }} />
        <Btn small onClick={commit}>추가</Btn>
        <Btn small onClick={() => { setAdding(false); setDraft(""); }}>취소</Btn>
      </div>
    );
  }
  return (
    <select value={value || ""} onChange={e => e.target.value === "__add__" ? setAdding(true) : onChange(e.target.value)}
      style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-md)", border: `1px solid ${C.line}`, background: C.panel, color: C.text, ...style }}>
      <option value="">{placeholder || "미분류"}</option>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
      <option value="__add__">+ 새 카테고리 추가</option>
    </select>
  );
}

/* ---------- 수치 토큰 렌더링 ---------- */
function formatMetric(m, fmt) {
  if (!m) return "?";
  let v;
  if (m.changeValue != null) v = `${m.changeValue}${m.unit}`;
  else if (m.beforeValue != null && m.afterValue != null) v = `${m.beforeValue}${m.unit}→${m.afterValue}${m.unit}`;
  else if (m.afterValue != null) v = `${m.afterValue}${m.unit}`;
  else v = "?";
  if (fmt === "rounded" && m.changeValue != null) v = `약 ${Math.round(m.changeValue / 5) * 5}${m.unit}`;
  return v;
}
function resolveTokenText(text, metrics) {
  if (!text) return "";
  return text.replace(/\{\{metric:([^|}]+)\|?([^}]*)\}\}/g, (_, id, fmt) => {
    const metric = metrics.find(x => x.id === id);
    return formatMetric(metric, fmt || "exact");
  });
}
function TokenText({ text, metrics }) {
  const parts = text.split(/(\{\{metric:[^}]+\}\})/g);
  return (
    <span>
      {parts.map((p, i) => {
        const m = p.match(/\{\{metric:([^|}]+)\|?([^}]*)\}\}/);
        if (!m) return <span key={i}>{p}</span>;
        const metric = metrics.find(x => x.id === m[1]);
        const cert = metric ? CERTAINTY[metric.certainty] : null;
        return (
          <span key={i} title={metric ? `${metric.metricName} · ${cert[0]} · 근거: ${metric.evidenceSource || "없음"}` : "삭제된 수치"}
            style={{ background: metric ? (metric.certainty === "verified" ? C.greenBg : C.orangeBg) : C.redBg,
              color: metric ? (metric.certainty === "verified" ? C.greenText : C.orangeText) : C.redText,
              padding: "var(--sp-1) var(--sp-2)", borderRadius: "var(--r-md)", fontWeight: 600, fontSize: "0.94em", cursor: "help" }}>
            {formatMetric(metric, m[2] || "exact")}
          </span>
        );
      })}
    </span>
  );
}

/* ---------- 공통 UI (와이어프레임 킷 톤 — 각진 박스, 아웃라인 태그) ---------- */
const Badge = ({ label, color, bg }) => (
  <span style={{ fontSize: "var(--fs-xs)", fontWeight: 700, color: "#fff", background: color,
    padding: "var(--sp-1) var(--sp-3)", borderRadius: "var(--r-xs)", display: "inline-block", lineHeight: 1.5,
    // 그리드 아이템이 되면 inline-block 이 block 으로 강제 변환(blockify)되어
    // 셀 전체 폭으로 늘어난다. 배지는 항상 내용 폭만 차지해야 한다.
    justifySelf: "start", width: "fit-content",
    // 호출부가 긴 문장을 label 로 넘기는 곳이 있다(이력서의 역량 목록 등).
    // nowrap 이면 그런 배지가 컨테이너를 밀어내 레이아웃 뷰포트까지 넓어진다.
    // keep-all 이라 한글 단어 중간에서는 끊기지 않고 띄어쓰기에서만 줄이 바뀐다.
    maxWidth: "100%", whiteSpace: "normal", wordBreak: "keep-all" }}>{label}</span>
);
const Card = ({ children, style, onClick, className }) => {
  // clickableProps 도 className("ui-click")을 넣으므로 호출부가 준 값과 합친다.
  // 합치지 않으면 뒤에 오는 쪽이 앞을 덮어써서, 예를 들어 반응형 클래스가
  // 조용히 사라지고 좁은 화면에서 그리드가 접히지 않는다.
  const interactive = onClick ? clickableProps(onClick) : {};
  const cls = [interactive.className, className].filter(Boolean).join(" ") || undefined;
  return (
  <div {...interactive} className={cls} style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: "var(--r-sm)", padding: "var(--sp-6)",
    boxShadow: "0 0 1px rgba(0,0,21,.08), 0 1px 3px rgba(0,0,21,.06)",
    cursor: onClick ? "pointer" : "default", transition: "border-color .15s, box-shadow .15s", ...style }}
    onMouseEnter={e => onClick && (e.currentTarget.style.borderColor = C.primary)}
    onMouseLeave={e => onClick && (e.currentTarget.style.borderColor = C.line)}>
    {children}
  </div>
  );
};
const H2 = ({ children }) => <h2 style={{ fontSize: "var(--fs-xl)", fontWeight: 700, margin: "0 0 14px", color: C.text }}>{children}</h2>;
const Label = ({ children }) => <div style={{ fontSize: "var(--fs-sm)", fontWeight: 600, color: C.faintText, marginBottom: 4, letterSpacing: ".02em" }}>{children}</div>;
const Btn = ({ children, primary, small, onClick, disabled, style, title, className }) => (
  <button onClick={onClick} disabled={disabled} title={title}
    className={["ui-click", className].filter(Boolean).join(" ")} style={{
    fontFamily: font,
    // 이 두 값은 타입/간격 스케일을 쓰지 않고 숫자로 남아 있었다.
    // 치환 스크립트가 삼항 표현식을 건너뛰었기 때문이다.
    fontSize: small ? "var(--fs-xs)" : "var(--fs-sm)", fontWeight: 600,
    padding: small ? "var(--sp-2) var(--sp-4)" : "var(--sp-3) var(--sp-6)",
    // 버튼 라벨은 절대 줄바꿈하지 않는다. "대화 초기화"가 "대화 초 / 기화"로
    // 쪼개지면 읽기 어렵다. 줄을 바꿔야 할 때는 버튼 통째로 넘어가야 하며,
    // 그건 부모 행의 .wrap-sm 이 처리한다.
    whiteSpace: "nowrap",
    borderRadius: "var(--r-sm)", border: primary ? `1px solid ${C.primary}` : `1px solid ${C.line}`, cursor: disabled ? "default" : "pointer",
    background: disabled ? C.lineSoft : primary ? C.primary : C.panel, color: disabled ? C.faintText : primary ? "#fff" : C.text, ...style }}>
    {children}
  </button>
);
// 클릭 가능하지만 <button> 으로 바꿀 수 없는 요소(내부에 블록 요소가 있거나
// 클릭 요소가 중첩되어 HTML 콘텐츠 모델을 위반하는 경우)에 키보드 접근성을 부여한다.
// role/tabIndex/onKeyDown 을 함께 주어 스크린리더에는 버튼으로 노출되고,
// Enter·Space 로 동작한다. display 가 바뀌지 않으므로 레이아웃은 그대로다.
const clickableProps = (onClick, { disabled = false, label } = {}) => ({
  role: "button",
  tabIndex: disabled ? -1 : 0,
  "aria-disabled": disabled || undefined,
  "aria-label": label,
  className: "ui-click",
  onClick: disabled ? undefined : onClick,
  onKeyDown: disabled ? undefined : (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onClick?.(e); }
  },
});

const Input = props => (
  <input {...props} style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-5)", borderRadius: "var(--r-sm)", border: `1px solid ${C.line}`,
    width: "100%", boxSizing: "border-box", background: C.panel, color: C.text, outline: "none", ...props.style }} />
);
const Textarea = props => (
  <textarea {...props} style={{ fontFamily: font, fontSize: "var(--fs-base)", lineHeight: 1.6, padding: "var(--sp-4) var(--sp-5)", borderRadius: "var(--r-sm)",
    border: `1px solid ${C.line}`, width: "100%", boxSizing: "border-box", background: C.panel, color: C.text, outline: "none",
    resize: "vertical", minHeight: 84, ...props.style }} />
);

/* ---------- 자동 저장 표시 ---------- */
function useAutosave(dep) {
  const [state, setState] = useState("saved");
  const first = useRef(true);
  useEffect(() => {
    if (first.current) { first.current = false; return; }
    setState("saving");
    const t = setTimeout(() => setState("saved"), 800);
    return () => clearTimeout(t);
  }, [dep]);
  return state;
}
const AutosaveIndicator = ({ state }) => (
  <span style={{ fontSize: "var(--fs-sm)", color: state === "saving" ? C.blueText : C.faintText, display: "inline-flex", alignItems: "center", gap: 5 }}>
    <span style={{ width: 6, height: 6, borderRadius: "var(--r-full)", background: state === "saving" ? C.blue : C.green }} />
    {state === "saving" ? "저장 중…" : "자동 저장됨"}
  </span>
);

/* ============================================================ APP */
/* ---------- 로컬 저장(localStorage) 지속성 훅 ---------- */
const STORAGE_PREFIX = "careeros:";
// 열려 있는 오버레이를 Escape 로 닫는다. 모달은 키보드만으로도 빠져나올 수 있어야 한다.
function useEscapeKey(active, onClose) {
  useEffect(() => {
    if (!active) return;
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [active, onClose]);
}
/* ---------- 토스트 ----------
   alert() 는 브라우저를 멈추고, 스타일을 입힐 수 없고, 여러 건을 겹쳐 보여줄 수도 없다.
   화면 좌하단(챗봇 FAB 반대편)에 쌓였다가 사라지는 비차단 알림으로 대체한다.
   컴포넌트 트리 어디서든 toast() 로 호출할 수 있도록 모듈 수준 구독 방식을 쓴다. */
const toastListeners = new Set();
let toastSeq = 0;
const toast = (message, tone = "info") => {
  toastListeners.forEach((fn) => fn({ id: ++toastSeq, message, tone }));
};
function Toaster() {
  const [items, setItems] = useState([]);
  useEffect(() => {
    const add = (t) => {
      setItems((prev) => [...prev, t]);
      // 오류는 읽을 시간이 더 필요하다
      setTimeout(() => setItems((prev) => prev.filter((i) => i.id !== t.id)), t.tone === "error" ? 8000 : 4500);
    };
    toastListeners.add(add);
    return () => { toastListeners.delete(add); };
  }, []);
  if (!items.length) return null;
  return (
    <div style={{ position: "fixed", left: 20, bottom: 20, zIndex: 70, display: "flex", flexDirection: "column", gap: 8, maxWidth: "min(400px, calc(100vw - 40px))" }}>
      {items.map((t) => (
        <div key={t.id}
          role={t.tone === "error" ? "alert" : "status"}
          aria-live={t.tone === "error" ? "assertive" : "polite"}
          style={{
            display: "flex", alignItems: "flex-start", gap: 8,
            background: C.panel, color: C.text,
            border: `1px solid ${t.tone === "error" ? C.red : C.line}`,
            borderLeft: `3px solid ${t.tone === "error" ? C.red : C.green}`,
            borderRadius: "var(--r-sm)", padding: "var(--sp-4) var(--sp-5)", fontSize: "var(--fs-base)", lineHeight: 1.5,
            boxShadow: "0 4px 14px rgba(0,0,21,.14)",
          }}>
          <span style={{ flex: 1, minWidth: 0 }}>{t.message}</span>
          <span {...clickableProps(() => setItems((prev) => prev.filter((i) => i.id !== t.id)), { label: "알림 닫기" })}
            style={{ cursor: "pointer", color: C.faintText, flexShrink: 0, lineHeight: 1 }}>
            <CIcon icon={cilX} width={12} height={12} aria-hidden="true" />
          </span>
        </div>
      ))}
    </div>
  );
}

function useIsMobile(breakpoint = 820) {
  const [isMobile, setIsMobile] = useState(() => typeof window !== "undefined" && window.innerWidth < breakpoint);
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < breakpoint);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, [breakpoint]);
  return isMobile;
}

/* ============================================================ 클라우드 동기화 (Supabase) ============================================================ */
// 동적 import: Claude.ai 아티팩트 미리보기 등 @supabase/supabase-js가 없는 환경에서도
// 나머지 Career OS 기능이 깨지지 않도록 방어. 배포된 사이트(Vercel + 환경변수)에서만 실제로 연결된다.
let _cloudSupabasePromise = null;
function getCloudClient() {
  if (_cloudSupabasePromise) return _cloudSupabasePromise;
  _cloudSupabasePromise = (async () => {
    let url, anonKey;
    try {
      url = typeof import.meta !== "undefined" ? import.meta.env?.VITE_SUPABASE_URL : undefined;
      anonKey = typeof import.meta !== "undefined" ? import.meta.env?.VITE_SUPABASE_ANON_KEY : undefined;
    } catch { /* import.meta 미지원 환경 */ }
    if (!url || !anonKey) return null;
    try {
      const mod = await import("@supabase/supabase-js");
      return mod.createClient(url, anonKey);
    } catch {
      return null;
    }
  })();
  return _cloudSupabasePromise;
}

let _cloudAuthPromise = null;
let _cachedCloudUser = null;
function ensureCloudAuth(supabase) {
  // 더 이상 익명 로그인을 하지 않는다. 로그인 안 한 상태는 오류가 아니라 정상적인 "로컬 전용" 상태다.
  if (!supabase) return Promise.resolve({ user: null, error: null });
  if (_cloudAuthPromise) return _cloudAuthPromise;
  _cloudAuthPromise = (async () => {
    try {
      const { data: { session }, error: sessErr } = await supabase.auth.getSession();
      if (sessErr) return { user: null, error: `세션 확인 실패: ${sessErr.message}` };
      _cachedCloudUser = session?.user || null;
      return { user: _cachedCloudUser, error: null };
    } catch (e) {
      return { user: null, error: `연결 중 예외 발생: ${e.message || String(e)}` };
    }
  })();
  return _cloudAuthPromise;
}
function resetCloudAuth() { _cloudAuthPromise = null; }
function getCachedCloudUser() { return _cachedCloudUser; }

async function signInWithGoogle(supabase) {
  if (!supabase) return { error: "Supabase 클라이언트가 없습니다." };
  const { error } = await supabase.auth.signInWithOAuth({
    provider: "google",
    options: { redirectTo: window.location.origin + window.location.pathname },
  });
  return { error: error ? error.message : null };
}

async function signOutCloud(supabase) {
  if (!supabase) return;
  try { await supabase.auth.signOut(); } catch { /* ignore */ }
  _cachedCloudUser = null;
  resetCloudAuth();
  window.location.reload();
}

/* ---------- 로그인 시점 로컬/클라우드 데이터 충돌 조정 ---------- */
let _reconcileResolve = null;
let _reconcileGate = new Promise(res => { _reconcileResolve = res; });
function resolveReconcile(mode) { if (_reconcileResolve) { _reconcileResolve(mode); _reconcileResolve = null; } }

const RECONCILE_KEYS = ["experiences", "metrics", "outputs", "applications", "skills", "certs", "awards", "resumeProfile",
  "masterEssays", "masterInterviews", "interviewCategories", "expCategories", "questionBlocks", "timelineActivities", "trash"];

function isMeaningfulValue(key, val) {
  if (val == null) return false;
  if (Array.isArray(val)) return val.length > 0;
  if (key === "resumeProfile") return !!(val.name || val.headline || val.targetRole);
  if (typeof val === "object") return Object.keys(val).length > 0;
  return false;
}
function summarizeValue(key, val) {
  if (Array.isArray(val)) return `${val.length}개`;
  if (key === "resumeProfile") return val.name || val.headline || "(내용 있음)";
  return "있음";
}
const RECONCILE_LABEL = { experiences: "경험", metrics: "성과 수치", outputs: "활용 문장", applications: "지원 현황", skills: "역량·스킬",
  certs: "자격증", awards: "수상기록", resumeProfile: "기본 이력서", masterEssays: "마스터 자소서",
  masterInterviews: "마스터 면접", interviewCategories: "면접 카테고리", expCategories: "경험 카테고리",
  questionBlocks: "질문별 블록", timelineActivities: "타임라인 활동", trash: "휴지통" };

async function reconcileOnLogin(supabase, user) {
  const local = {};
  RECONCILE_KEYS.forEach(k => {
    try { const raw = window.localStorage.getItem(STORAGE_PREFIX + k); if (raw != null) local[k] = JSON.parse(raw); } catch { /* ignore */ }
  });
  let cloud = {};
  try {
    const { data, error } = await supabase.from("career_os_state").select("key, value").eq("user_id", user.id);
    if (!error && data) data.forEach(row => { cloud[row.key] = row.value; });
  } catch { /* ignore */ }

  const diffs = [];
  RECONCILE_KEYS.forEach(k => {
    const l = local[k], c = cloud[k];
    if (isMeaningfulValue(k, l) && isMeaningfulValue(k, c) && JSON.stringify(l) !== JSON.stringify(c)) {
      diffs.push({ key: k, label: RECONCILE_LABEL[k] || k, localSummary: summarizeValue(k, l), cloudSummary: summarizeValue(k, c) });
    }
  });
  return { diffs };
}


function usePersisted(key, initialValue) {
  const [state, setState] = useState(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_PREFIX + key);
      if (raw != null) return JSON.parse(raw);
    } catch (e) { /* 저장소 접근 불가 시 기본값으로 진행 */ }
    return initialValue;
  });
  const [cloudStatus, setCloudStatus] = useState("idle"); // idle | syncing | synced | offline | error
  const lastSyncedRef = useRef(null);
  const cloudReadyRef = useRef(false);

  // 1) 로컬 저장 — 항상 즉시 (Supabase 연결 여부와 무관하게 안전망 역할)
  useEffect(() => {
    try { window.localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(state)); }
    catch (e) { /* 저장 실패해도 앱은 계속 동작 */ }
  }, [key, state]);

  // 2) 클라우드 동기화 — 로그인 시점 충돌 조정(reconcile) 게이트를 먼저 기다린다
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const supabase = await getCloudClient();
      if (!supabase) { setCloudStatus("offline"); return; }
      const { user } = await ensureCloudAuth(supabase);
      if (!user || cancelled) { setCloudStatus("offline"); return; } // 로그인 안 함 = 정상적인 로컬 전용 상태

      const mode = await _reconcileGate; // "local" | "cloud" | "none" — 로그인 시점에 1회 결정됨
      if (cancelled) return;
      try {
        if (mode === "local") {
          await supabase.from("career_os_state").upsert({ user_id: user.id, key, value: state });
          lastSyncedRef.current = JSON.stringify(state);
        } else {
          const { data, error } = await supabase.from("career_os_state")
            .select("value").eq("user_id", user.id).eq("key", key).maybeSingle();
          if (cancelled) return;
          if (error) throw error;
          if (data && data.value !== null && data.value !== undefined) {
            setState(data.value);
            lastSyncedRef.current = JSON.stringify(data.value);
          } else {
            await supabase.from("career_os_state").upsert({ user_id: user.id, key, value: state });
            lastSyncedRef.current = JSON.stringify(state);
          }
        }
        cloudReadyRef.current = true;
        setCloudStatus("synced");
      } catch (e) {
        console.error("[cloud sync 실패]", key, e);
        setCloudStatus("error");
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key]);

  // 3) 이후 변경분을 클라우드로 반영 (디바운스)
  useEffect(() => {
    if (!cloudReadyRef.current) return;
    const serialized = JSON.stringify(state);
    if (serialized === lastSyncedRef.current) return;
    setCloudStatus("syncing");
    const t = setTimeout(async () => {
      const supabase = await getCloudClient();
      if (!supabase) return;
      const { user } = await ensureCloudAuth(supabase);
      if (!user) return;
      try {
        await supabase.from("career_os_state").upsert({ user_id: user.id, key, value: state });
        lastSyncedRef.current = serialized;
        setCloudStatus("synced");
      } catch (e) {
        console.error("[cloud sync 실패]", key, e);
        setCloudStatus("error");
      }
    }, 800);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state]);

  return [state, setState, cloudStatus];
}

function App() {
  const [nav, setNav] = useState("home"); // home | analyze | archive | apply | resume
  const backupInputRef = useRef(null);
  const [experiences, setExperiences, cloudStatus] = usePersisted("experiences", []);
  const [metrics, setMetrics] = usePersisted("metrics", []);
  const [outputs, setOutputs] = usePersisted("outputs", []);
  const [applications, setApplications] = usePersisted("applications", []);
  const [skills, setSkills] = usePersisted("skills", []);
  const [certs, setCerts] = usePersisted("certs", []);
  const [awards, setAwards] = usePersisted("awards", []);
  const [resumeProfile, setResumeProfile] = usePersisted("resumeProfile", { name: "", targetRole: "", headline: "", phone: "", email: "" });
  const [detailId, setDetailId] = useState(null);     // 경험 상세
  const [appDetailId, setAppDetailId] = useState(null); // 지원 상세
  const [analyzeId, setAnalyzeId] = useState(null);   // 분석 중 경험
  const [trash, setTrash] = usePersisted("trash", []); // { id, type, label, deletedAt, payload }
  const [masterEssays, setMasterEssays] = usePersisted("masterEssays", seedMasterEssays);
  const [masterInterviews, setMasterInterviews] = usePersisted("masterInterviews", seedMasterInterviews);
  const [interviewCategories, setInterviewCategories] = usePersisted("interviewCategories", ["성과", "실패", "협업", "갈등", "인성"]);
  const [expCategories, setExpCategories] = usePersisted("expCategories", ["온라인 쇼핑몰 인턴", "동아리 활동"]);
  const [questionBlocks, setQuestionBlocks] = usePersisted("questionBlocks", emptyQuestionBlocks);
  const [timelineActivities, setTimelineActivities] = usePersisted("timelineActivities", []);
  const [reviewChatHistory, setReviewChatHistory] = usePersisted("reviewChatHistory", []);
  const [personalChatHistory, setPersonalChatHistory] = usePersisted("personalChatHistory", []);
  const addInterviewCategory = (c) => setInterviewCategories(prev => prev.includes(c) ? prev : [...prev, c]);
  const addExpCategory = (c) => setExpCategories(prev => prev.includes(c) ? prev : [...prev, c]);

  const [authUser, setAuthUser] = useState(undefined); // undefined=확인 중, null=로그아웃, object=로그인됨
  const [authSupabase, setAuthSupabase] = useState(null);
  const [conflict, setConflict] = useState(null); // { diffs } — 로컬/클라우드 둘 다 의미 있는 데이터가 있어 다를 때만 표시
  const [authLoading, setAuthLoading] = useState(false);

  useEffect(() => {
    (async () => {
      const supabase = await getCloudClient();
      setAuthSupabase(supabase);
      if (!supabase) { resolveReconcile("none"); setAuthUser(null); return; }
      const { user } = await ensureCloudAuth(supabase);
      setAuthUser(user);
      if (!user) { resolveReconcile("none"); return; }
      try {
        const result = await reconcileOnLogin(supabase, user);
        if (result.diffs.length > 0) setConflict(result);
        else resolveReconcile("none");
      } catch {
        resolveReconcile("none");
      }
      if (window.location.hash.includes("access_token") || window.location.search.includes("code=")) {
        window.history.replaceState(null, "", window.location.pathname);
      }
    })();
  }, []);

  const handleGoogleLogin = async () => {
    setAuthLoading(true);
    const sb = authSupabase || await getCloudClient();
    if (!sb) { setAuthLoading(false); toast("Supabase가 설정되어 있지 않습니다.", "error"); return; }
    const { error } = await signInWithGoogle(sb);
    if (error) { setAuthLoading(false); toast("로그인에 실패했습니다. " + error, "error"); }
  };
  const handleLogout = async () => { if (authSupabase) await signOutCloud(authSupabase); };

  const resolveConflict = (mode) => { resolveReconcile(mode); setConflict(null); };

  const isBlankSlate = experiences.length === 0 && applications.length === 0 && skills.length === 0 && certs.length === 0;
  const loadDemoData = () => {
    setExperiences(seedExperiences);
    setMetrics(seedMetrics);
    setOutputs(seedOutputs);
    setApplications(seedApplications);
    setSkills(seedSkills);
    setCerts(seedCerts);
    setQuestionBlocks(seedQuestionBlocks);
  };

  const exportBackup = () => {
    const payload = {
      _type: "career-os-backup", _version: 1, exportedAt: new Date().toISOString(),
      experiences, metrics, outputs, applications, skills, certs, awards,
      resumeProfile, masterEssays, masterInterviews, interviewCategories,
      expCategories, questionBlocks,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `career-os-백업-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  };

  const importBackup = (file) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const d = JSON.parse(reader.result);
        if (d._type !== "career-os-backup") throw new Error("Career OS 백업 파일이 아닙니다.");
        if (!window.confirm("불러오면 현재 데이터를 덮어씁니다. 계속할까요?")) return;
        if (d.experiences) setExperiences(d.experiences);
        if (d.metrics) setMetrics(d.metrics);
        if (d.outputs) setOutputs(d.outputs);
        if (d.applications) setApplications(d.applications);
        if (d.skills) setSkills(d.skills);
        if (d.certs) setCerts(d.certs);
        if (d.awards) setAwards(d.awards);
        if (d.resumeProfile) setResumeProfile(d.resumeProfile);
        if (d.masterEssays) setMasterEssays(d.masterEssays);
        if (d.masterInterviews) setMasterInterviews(d.masterInterviews);
        if (d.interviewCategories) setInterviewCategories(d.interviewCategories);
        if (d.expCategories) setExpCategories(d.expCategories);
        if (d.questionBlocks) setQuestionBlocks(d.questionBlocks);
        toast("백업을 불러왔습니다.");
      } catch (err) {
        toast("백업 파일을 읽지 못했습니다. " + err.message, "error");
      }
    };
    reader.readAsText(file);
  };

  const addTrash = (type, label, payload) => setTrash(prev => [
    { id: "t_" + Date.now() + Math.random().toString(36).slice(2, 6), type, label, deletedAt: new Date().toISOString(), payload },
    ...prev,
  ]);
  const restoreTrash = (id) => {
    const entry = trash.find(t => t.id === id);
    if (!entry) return;
    const { type, payload } = entry;
    if (type === "experience") setExperiences(prev => [...prev, payload]);
    else if (type === "skill") setSkills(prev => [...prev, payload]);
    else if (type === "cert") setCerts(prev => [...prev, payload]);
    else if (type === "award") setAwards(prev => [...prev, payload]);
    else if (type === "application") setApplications(prev => [...prev, payload]);
    else if (type === "requirement") setApplications(prev => prev.map(a => a.id === payload.appId ? { ...a, requirements: [...a.requirements, payload.item] } : a));
    else if (type === "essay") setApplications(prev => prev.map(a => a.id === payload.appId ? { ...a, essays: [...a.essays, payload.item] } : a));
    else if (type === "interview") setApplications(prev => prev.map(a => a.id === payload.appId ? { ...a, interviews: [...a.interviews, payload.item] } : a));
    else if (type === "timeline_activity") setTimelineActivities(prev => [...prev, payload]);
    setTrash(prev => prev.filter(t => t.id !== id));
  };
  const purgeTrash = (id) => setTrash(prev => prev.filter(t => t.id !== id));
  const clearTrash = () => setTrash([]);

  const go = (n) => { setNav(n); setDetailId(null); setAppDetailId(null); if (n !== "analyze") setAnalyzeId(null); };

  const openAnalyze = (id) => { setAnalyzeId(id); setNav("analyze"); setDetailId(null); };
  const openDetail = (id) => { setDetailId(id); setNav("archive"); };

  const TOP_NAV = [
    { key: "home", label: "홈", nav: "home" },
    { key: "myexp", label: "내 경험", subTabs: [["archive", "경험 보관함"], ["timeline", "타임라인"], ["skills", "역량·스킬"]] },
    { key: "prep", label: "지원 준비", subTabs: [["apply", "지원 관리"], ["master", "자소서·면접 준비"], ["resume", "기본 이력서"]] },
  ];
  const activeTop = TOP_NAV.find(t => t.key === nav || (t.subTabs && t.subTabs.some(([k]) => k === nav))) || TOP_NAV[0];
  const isMobile = useIsMobile();
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [showGuide, setShowGuide] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [chatOpen, setChatOpen] = useState(false);
  useEscapeKey(chatOpen, () => setChatOpen(false));
  useEscapeKey(showGuide, () => setShowGuide(false));
  useEscapeKey(showSettings, () => setShowSettings(false));

  return (
    <div className="min-h-screen" style={{ fontFamily: font, background: C.bg, display: "flex", flexDirection: "column", color: C.text }}>
      {/* 키보드 사용자가 사이드바를 건너뛰고 본문으로 바로 이동 */}
      <a href="#main-content" className="skip-link">본문으로 건너뛰기</a>
      <Toaster />
      {/* 상단 헤더 — 실제 CoreUI CHeader */}
      <CHeader position="sticky" className="mb-0" style={{ zIndex: 20 }}>
        <CContainer fluid className="d-flex justify-content-between align-items-center flex-wrap" style={{ gap: 10 }}>
          <div className="d-flex align-items-center flex-wrap" style={{ gap: 10 }}>
            {/* 마크는 22px 까지 헤더 높이를 늘리지 않는다 — 브랜드 줄 상자가 이미 30px 라
                그 안에 들어간다. 32px 부터 헤더가 커진다. */}
            <CHeaderBrand className="d-flex align-items-center" style={{ gap: 8, fontWeight: 800 }}>
              <BrandMark size={22} />
              <BrandWordmark size="inherit" />
            </CHeaderBrand>
            <span className="text-body-secondary">/</span>
            <span className="text-body-secondary fw-semibold">{activeTop.label}</span>
            {isMobile && (
              <CHeaderToggler onClick={() => setMobileMenuOpen(o => !o)}
                aria-expanded={mobileMenuOpen} aria-controls="main-nav"
                // CoreUI 기본값이 20px 이라 320px 화면에서 "메뉴 접기"가 두 줄로 쪼개졌다.
                style={{ fontSize: "var(--fs-base)", whiteSpace: "nowrap", padding: "var(--sp-1) var(--sp-2)" }}>
                {mobileMenuOpen ? "메뉴 접기 ▴" : "메뉴 ▾"}</CHeaderToggler>
            )}
          </div>
          <CHeaderNav className="d-flex align-items-center flex-wrap" style={{ gap: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 5, fontSize: "var(--fs-sm)", color: C.faintText }}>
              <span style={{ width: 6, height: 6, borderRadius: "var(--r-full)", background:
                cloudStatus === "synced" ? C.green : cloudStatus === "syncing" ? C.blue : cloudStatus === "error" ? C.red : C.faint }} />
              {cloudStatus === "synced" ? "클라우드에 저장됨" : cloudStatus === "syncing" ? "동기화 중…" : cloudStatus === "error" ? "동기화 실패 (로컬엔 저장됨)" : "저장됨 · 이 브라우저에만"}
            </div>
            <span {...clickableProps(exportBackup)} style={{ fontSize: "var(--fs-sm)", color: C.sub, cursor: "pointer", textDecoration: "underline" }}>백업 다운로드</span>
            {authUser === undefined ? null : authUser ? (
              <span {...clickableProps(handleLogout)} style={{ fontSize: "var(--fs-sm)", color: C.sub, cursor: "pointer" }}>
                {authUser.email || "로그인됨"} · <span style={{ textDecoration: "underline" }}>로그아웃</span>
              </span>
            ) : (
              <Btn small onClick={handleGoogleLogin} disabled={authLoading}>{authLoading ? "이동 중…" : "Google로 로그인"}</Btn>
            )}
            <span {...clickableProps(() => setShowGuide(true))} title="사용 가이드" style={{ cursor: "pointer", fontSize: "var(--fs-md)", color: C.sub, width: 26, height: 26, borderRadius: "50%", border: `1px solid ${C.line}`, display: "flex", alignItems: "center", justifyContent: "center" }}>?</span>
            <span {...clickableProps(() => setShowSettings(true))} title="설정" style={{ cursor: "pointer", fontSize: "var(--fs-md)", color: C.sub, width: 26, height: 26, borderRadius: "50%", border: `1px solid ${C.line}`, display: "flex", alignItems: "center", justifyContent: "center" }}><CIcon icon={cilCog} width={14} height={14} aria-hidden="true" /></span>
          </CHeaderNav>
        </CContainer>
      </CHeader>

      <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", flex: 1, minHeight: 0 }}>
      {/* Sidebar — 실제 CoreUI CSidebar
          CoreUI 의 사이드바는 .show 클래스로 열고 닫는 오버레이 드로어다.
          이 앱은 대신 흐름 안에서 펼쳐지는 인라인 메뉴를 쓰므로
          (position: static, width: 100%, height: auto) 드로어 로직을
          쓰지 않는다. visible prop 은 CoreUI 내부에서 mobile 판정 직후
          강제로 꺼지므로 여기서는 쓸 수 없고, 숨김 마진만 직접 무효화한다. */}
      {(!isMobile || mobileMenuOpen) && (
        <CSidebar id="main-nav" colorScheme="dark" aria-label="주 메뉴"
          style={{ width: isMobile ? "100%" : 220, position: isMobile ? "static" : "sticky", top: 53,
            height: isMobile ? "auto" : "calc(100vh - 53px)",
            // CoreUI 의 .sidebar:not(.show) 가 margin-inline-start: -16rem 으로
            // 사이드바를 화면 밖에 숨겨 둔다. 이 앱은 오버레이 드로어가 아니라
            // 흐름 안에서 펼쳐지는 인라인 메뉴를 쓰므로 그 숨김을 무효화한다.
            marginInlineStart: 0,
            // CoreUI 는 .sidebar 에 flex: 0 0 16rem 을 건다. 모바일에서는 부모가
            // 세로 flex 라 flex-basis 가 '높이'를 결정하므로 위의 height: auto 가
            // 무시되고 항목 아래로 빈 공간이 256px 까지 남는다. 내용에 맞게 접는다.
            flex: isMobile ? "0 0 auto" : undefined }}>
          <CSidebarNav>
            {TOP_NAV.map(t => (
              t.subTabs ? (
                <CNavGroup key={t.key} toggler={t.label} visible={activeTop.key === t.key}
                  onClick={() => { if (activeTop.key !== t.key) go(t.subTabs[0][0]); }}>
                  {t.subTabs.map(([k, l]) => (
                    <CNavItem key={k} href="#" active={nav === k} aria-current={nav === k ? "page" : undefined}
                      onClick={e => { e.preventDefault(); go(k); if (isMobile) setMobileMenuOpen(false); }}>{l}</CNavItem>
                  ))}
                </CNavGroup>
              ) : (
                <CNavItem key={t.key} href="#" active={activeTop.key === t.key} aria-current={activeTop.key === t.key ? "page" : undefined}
                  onClick={e => { e.preventDefault(); go(t.nav); if (isMobile) setMobileMenuOpen(false); }}>{t.label}</CNavItem>
              )
            ))}
          </CSidebarNav>
        </CSidebar>
      )}

      {/* Main */}
      <main id="main-content" style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        <div className="page-body" style={{ padding: isMobile ? "var(--sp-6)" : "var(--sp-8) var(--sp-8)", maxWidth: 1120, minWidth: 0, width: "100%", boxSizing: "border-box" }}>
        {nav === "home" && <Home experiences={experiences} applications={applications} onGoAnalyze={() => go("analyze")} onGoImport={() => go("import")} onOpenDetail={openDetail} onOpenApp={id => { setNav("apply"); setAppDetailId(id); }} isBlankSlate={isBlankSlate} onLoadDemo={loadDemoData} onGoGuide={() => setShowGuide(true)} />}
        {nav === "chat" && <PersonalAssistant experiences={experiences} skills={skills} certs={certs} awards={awards} resumeProfile={resumeProfile} applications={applications} metrics={metrics}
          history={personalChatHistory} setHistory={setPersonalChatHistory} onGo={go} />}
        {nav === "analyze" && <Analyze experiences={experiences} setExperiences={setExperiences} analyzeId={analyzeId} setAnalyzeId={setAnalyzeId} metrics={metrics} setMetrics={setMetrics} onDone={openDetail} />}
        {nav === "archive" && !detailId && <Archive experiences={experiences} setExperiences={setExperiences} metrics={metrics} setMetrics={setMetrics} outputs={outputs} setOutputs={setOutputs} onOpen={openDetail} onAnalyze={openAnalyze} onGoImport={() => go("import")} addTrash={addTrash} expCategories={expCategories} addExpCategory={addExpCategory} questionBlocks={questionBlocks} setQuestionBlocks={setQuestionBlocks} reviewChatHistory={reviewChatHistory} setReviewChatHistory={setReviewChatHistory} />}
        {nav === "archive" && detailId && <ExperienceDetail exp={experiences.find(e => e.id === detailId)} metrics={metrics} setMetrics={setMetrics} outputs={outputs} setOutputs={setOutputs} setExperiences={setExperiences} onBack={() => setDetailId(null)} onAnalyze={openAnalyze} onDeleted={() => setDetailId(null)} addTrash={addTrash} />}
        {nav === "timeline" && <Timeline experiences={experiences} setExperiences={setExperiences} activities={timelineActivities} setActivities={setTimelineActivities} addTrash={addTrash} onOpenExp={openDetail} onAnalyze={openAnalyze} onGoArchive={() => go("archive")} />}
        {nav === "import" && <ImportFlow setExperiences={setExperiences} setSkills={setSkills} setCerts={setCerts} setResumeProfile={setResumeProfile} onDone={openDetail} experiences={experiences} />}
        {nav === "skills" && <Skills skills={skills} setSkills={setSkills} experiences={experiences} onOpenExp={openDetail} addTrash={addTrash} />}
        {nav === "apply" && !appDetailId && <Applications applications={applications} setApplications={setApplications} onOpen={setAppDetailId} addTrash={addTrash} />}
        {nav === "apply" && appDetailId && <ApplicationDetail app={applications.find(a => a.id === appDetailId)} setApplications={setApplications} experiences={experiences} outputs={outputs} metrics={metrics} onBack={() => setAppDetailId(null)} onOpenExp={openDetail} addTrash={addTrash} interviewCategories={interviewCategories} addInterviewCategory={addInterviewCategory} />}
        {nav === "master" && <MasterPrep essays={masterEssays} setEssays={setMasterEssays} interviews={masterInterviews} setInterviews={setMasterInterviews} experiences={experiences} metrics={metrics} resumeProfile={resumeProfile} interviewCategories={interviewCategories} addInterviewCategory={addInterviewCategory} />}
        {nav === "resume" && <Resume experiences={experiences} outputs={outputs} metrics={metrics} resumeProfile={resumeProfile} setResumeProfile={setResumeProfile} skills={skills} certs={certs} setCerts={setCerts} awards={awards} setAwards={setAwards} addTrash={addTrash} />}
        {nav === "trash" && <Trash trash={trash} onRestore={restoreTrash} onPurge={purgeTrash} onClear={clearTrash} />}
        </div>
      </main>
      </div>

      {/* 플로팅 AI 물어보기 버튼 */}
      {!chatOpen && (
        <button onClick={() => setChatOpen(true)} title="AI에게 물어보기" aria-label="AI에게 물어보기"
          className="fab-safe" style={{
          // fab-safe 가 아이폰 홈 인디케이터만큼 더 띄운다. 아래 right/bottom 은
          // env() 를 모르는 브라우저를 위한 기본값이다.
          "--fab-inset": isMobile ? "16px" : "28px",
          position: "fixed", right: isMobile ? 16 : 28, bottom: isMobile ? 16 : 28, zIndex: 40,
          width: 52, height: 52, borderRadius: "var(--r-full)", background: C.primary, color: "#fff", border: "none",
          boxShadow: "0 4px 14px rgba(0,0,0,.18)", cursor: "pointer", fontSize: "var(--fs-2xl)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <CIcon icon={cilSpeech} width={22} height={22} aria-hidden="true" />
        </button>
      )}
      {chatOpen && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.35)", zIndex: 50, display: "flex", justifyContent: "flex-end", alignItems: isMobile ? "stretch" : "flex-end", padding: isMobile ? 0 : "var(--sp-7)" }}
          onClick={(e) => { if (e.target === e.currentTarget) setChatOpen(false); }}>
          <div role="dialog" aria-modal="true" aria-label="AI 어시스턴트" style={{ width: isMobile ? "100%" : 420, maxHeight: isMobile ? "100%" : "80vh", height: isMobile ? "100%" : "auto",
            background: C.bg, borderRadius: isMobile ? 0 : "var(--r-xl)", overflowY: "auto", padding: "var(--sp-7)", boxShadow: "0 8px 30px rgba(0,0,0,.2)" }}>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 4 }}>
              <span {...clickableProps(() => setChatOpen(false), { label: "닫기" })} style={{ cursor: "pointer", fontSize: "var(--fs-xl)", color: C.faintText }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
            </div>
            <PersonalAssistant experiences={experiences} skills={skills} certs={certs} awards={awards} resumeProfile={resumeProfile} applications={applications} metrics={metrics}
              history={personalChatHistory} setHistory={setPersonalChatHistory} onGo={() => setChatOpen(false)} />
          </div>
        </div>
      )}

      {/* 가이드 모달 */}
      {showGuide && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.35)", zIndex: 50, display: "flex", justifyContent: "center", alignItems: "flex-start", padding: isMobile ? 0 : "calc(var(--sp-8) + var(--sp-5)) var(--sp-7)", overflowY: "auto" }}
          onClick={(e) => { if (e.target === e.currentTarget) setShowGuide(false); }}>
          <div role="dialog" aria-modal="true" aria-label="사용 가이드" className={isMobile ? "modal-full-mobile" : undefined} style={{ width: "100%", maxWidth: 920, background: C.bg, borderRadius: isMobile ? 0 : "var(--r-xl)", padding: isMobile ? "var(--sp-6)" : "var(--sp-8)", boxShadow: "0 8px 30px rgba(0,0,0,.2)" }}>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 4 }}>
              <span {...clickableProps(() => setShowGuide(false), { label: "닫기" })} style={{ cursor: "pointer", fontSize: "var(--fs-xl)", color: C.faintText }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
            </div>
            <Guide onGo={(n) => { setShowGuide(false); go(n); }} />
          </div>
        </div>
      )}

      {/* 설정 모달: 백업 불러오기 · 전체 초기화 · 휴지통 */}
      {showSettings && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.35)", zIndex: 50, display: "flex", justifyContent: "center", alignItems: "flex-start", padding: isMobile ? 0 : "calc(var(--sp-8) + var(--sp-5)) var(--sp-7)", overflowY: "auto" }}
          onClick={(e) => { if (e.target === e.currentTarget) setShowSettings(false); }}>
          <div role="dialog" aria-modal="true" aria-label="설정" className={isMobile ? "modal-full-mobile" : undefined} style={{ width: "100%", maxWidth: 640, background: C.bg, borderRadius: isMobile ? 0 : "var(--r-xl)", padding: isMobile ? "var(--sp-6)" : "var(--sp-8)", boxShadow: "0 8px 30px rgba(0,0,0,.2)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <H2>설정</H2>
              <span {...clickableProps(() => setShowSettings(false), { label: "닫기" })} style={{ cursor: "pointer", fontSize: "var(--fs-xl)", color: C.faintText }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
            </div>

            <Card style={{ marginBottom: 16 }}>
              <Label>백업</Label>
              <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }} className="wrap-sm">
                <Btn small onClick={exportBackup}>백업 다운로드</Btn>
                <Btn small onClick={() => backupInputRef.current?.click()}>백업 불러오기</Btn>
                <input ref={backupInputRef} type="file" accept="application/json" style={{ display: "none" }}
                  onChange={e => { const f = e.target.files[0]; if (f) importBackup(f); e.target.value = ""; }} />
              </div>
            </Card>

            <Card style={{ marginBottom: 16 }}>
              <Label>휴지통 ({trash.length})</Label>
              <div style={{ marginTop: 8 }}>
                <Trash trash={trash} onRestore={restoreTrash} onPurge={purgeTrash} onClear={clearTrash} />
              </div>
            </Card>

            <Card>
              <Label>초기화</Label>
              <div style={{ marginTop: 8 }}>
                <span {...clickableProps(async () => {
                  if (!window.confirm("저장된 모든 데이터를 지우고 초기 상태로 되돌릴까요? (클라우드에 저장된 데이터도 함께 지워집니다) 되돌릴 수 없습니다.")) return;
                  Object.keys(window.localStorage).filter(k => k.startsWith(STORAGE_PREFIX)).forEach(k => window.localStorage.removeItem(k));
                  try {
                    const supabase = await getCloudClient();
                    if (supabase) {
                      const { user } = await ensureCloudAuth(supabase);
                      if (user) await supabase.from("career_os_state").delete().eq("user_id", user.id);
                    }
                  } catch (e) { console.error("[클라우드 초기화 실패]", e); }
                  window.location.reload();
                })} style={{ fontSize: "var(--fs-sm)", color: C.redText, cursor: "pointer", textDecoration: "underline" }}>
                  전체 데이터 초기화
                </span>
              </div>
            </Card>
          </div>
        </div>
      )}

      {/* 로그인 시 로컬/클라우드 데이터 충돌 — 절대 조용히 덮어쓰지 않는다 */}
      {conflict && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.5)", zIndex: 60, display: "flex", justifyContent: "center", alignItems: "center", padding: "var(--sp-7)" }} className="wrap-sm">
          <div style={{ width: "100%", maxWidth: 560, background: C.bg, borderRadius: "var(--r-xl)", padding: "var(--sp-8)", boxShadow: "0 8px 30px rgba(0,0,0,.3)" }}>
            <H2>어느 데이터를 사용할까요?</H2>
            <div style={{ fontSize: "var(--fs-base)", color: C.sub, lineHeight: 1.65, marginBottom: 16 }}>
              이 브라우저와 클라우드(계정) 양쪽에 서로 다른 데이터가 있습니다. 실수로 자소서 등 작성한 내용이 사라지지 않도록, 어느 쪽을 남길지 직접 선택해야 합니다. <b>선택한 쪽이 다른 쪽을 덮어씁니다.</b>
            </div>
            <div style={{ marginBottom: 18, maxHeight: 220, overflowY: "auto" }}>
              {conflict.diffs.map(d => (
                <div key={d.key} style={{ display: "flex", justifyContent: "space-between", padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }}>
                  <span style={{ fontWeight: 700 }}>{d.label}</span>
                  <span style={{ color: C.sub }}>이 브라우저 {d.localSummary} · 클라우드 {d.cloudSummary}</span>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 10 }} className="wrap-sm">
              <Btn style={{ flex: 1 }} onClick={() => resolveConflict("local")}>이 브라우저 데이터 사용</Btn>
              <Btn primary style={{ flex: 1 }} onClick={() => resolveConflict("cloud")}>클라우드 데이터 사용</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================================================ 랜딩(히어로) */
function Landing({ onStart, onGoogle }) {
  const [googleLoading, setGoogleLoading] = useState(false);
  const steps = [
    { n: 1, title: "경험을 넣으세요", desc: "자소서·이력서 파일을 첨부하거나, 그냥 아무거나 적어보세요. AI가 구조화해드립니다." },
    { n: 2, title: "AI가 정리·검토합니다", desc: "성과 수치, 역량, 빠진 정보를 짚어주고, 자소서·면접 문장을 함께 다듬습니다." },
    { n: 3, title: "지원 준비를 관리하세요", desc: "회사·직무별로 요구 역량 매칭부터 최종 이력서까지 한곳에서." },
  ];
  const handleGoogle = async () => {
    setGoogleLoading(true);
    try {
      const sb = await getCloudClient();
      if (sb) { await signInWithGoogle(sb); return; } // 로그인 후 redirect로 돌아옴
    } catch { /* ignore */ }
    setGoogleLoading(false);
    onGoogle();
  };
  return (
    <div className="min-h-screen" style={{ fontFamily: font, background: C.bg, color: C.text }}>
      <div style={{ maxWidth: 880, margin: "0 auto", padding: "60px var(--sp-7) 80px" }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10, marginBottom: 18 }}>
            <BrandMark size={44} />
            <BrandWordmark size="var(--fs-xl)" />
          </div>
          <h1 style={{ fontSize: "var(--fs-5xl)", fontWeight: 800, margin: "0 0 14px", lineHeight: 1.35 }}>
            흩어진 경험을, 이력서·자소서로 바로 쓸 수 있게
          </h1>
          <div style={{ fontSize: "var(--fs-md)", color: C.sub, lineHeight: 1.6, marginBottom: 28 }}>
            자소서·이력서 파일을 넣으면 AI가 경험을 정리하고, 지원 준비까지 한곳에서 관리합니다.
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "center", flexWrap: "wrap" }} className="wrap-sm">
            <Btn primary onClick={onStart} style={{ padding: "var(--sp-5) var(--sp-8)", fontSize: "var(--fs-md)" }}>바로 시작하기</Btn>
            <Btn onClick={handleGoogle} disabled={googleLoading} style={{ padding: "var(--sp-5) var(--sp-8)", fontSize: "var(--fs-md)" }}>
              {googleLoading ? "이동 중…" : "구글로 계속하기"}
            </Btn>
          </div>
        </div>

        <div style={{ marginBottom: 48, borderRadius: "var(--r-xl)", border: `1px solid ${C.line}`, background: C.panel, boxShadow: "0 10px 40px rgba(0,0,0,.06)", overflow: "hidden" }}>
          <img src="/landing-screenshot.png" alt="Career OS 화면 예시" style={{ width: "100%", display: "block" }}
            onError={e => { e.target.style.display = "none"; e.target.nextSibling.style.display = "flex"; }} />
          <div style={{ display: "none", height: 360, alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 10, background: C.accent, color: C.faintText }}>
            <Icon name="layers" size={40} />
            <div style={{ fontSize: "var(--fs-base)" }}>화면 미리보기</div>
          </div>
        </div>

        <Card>
          <Label>작동 방식</Label>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginTop: 10 }} className="stack-sm">
            {steps.map(s => (
              <div key={s.n} style={{ padding: "var(--sp-6) var(--sp-2)" }}>
                <div style={{ fontSize: "var(--fs-3xl)", fontWeight: 800, color: C.primary, marginBottom: 8 }}>{s.n}</div>
                <div style={{ fontSize: "var(--fs-md)", fontWeight: 700, marginBottom: 6 }}>{s.title}</div>
                <div style={{ fontSize: "var(--fs-sm)", color: C.sub, lineHeight: 1.55 }}>{s.desc}</div>
              </div>
            ))}
          </div>
        </Card>

        <div style={{ textAlign: "center", marginTop: 32 }}>
          <Btn primary onClick={onStart} style={{ padding: "var(--sp-5) var(--sp-8)", fontSize: "var(--fs-md)" }}>바로 시작하기</Btn>
        </div>
      </div>
    </div>
  );
}

/* ============================================================ 라우팅 루트 — 첫 방문자는 랜딩, 재방문자는 /app으로 */
export default function Root() {
  useSpoqaHanSansFont();
  const [screen, setScreen] = useState(() => {
    try {
      if (window.location.pathname.startsWith("/app")) return "app";
      const hasData = RECONCILE_KEYS.some(k => {
        try { return !!window.localStorage.getItem(STORAGE_PREFIX + k); } catch { return false; }
      });
      if (hasData) return "app";
      return "landing";
    } catch { return "app"; }
  });

  useEffect(() => {
    try {
      if (screen === "app" && window.location.pathname !== "/app") window.history.replaceState(null, "", "/app");
      if (screen === "landing" && window.location.pathname !== "/") window.history.replaceState(null, "", "/");
    } catch { /* 샌드박스 미리보기 등에서는 무시 */ }
  }, [screen]);

  if (screen === "landing") {
    return <Landing onStart={() => setScreen("app")} onGoogle={() => setScreen("app")} />;
  }
  return <App />;
}

/* ============================================================ 홈 */
/* ============================================================ 브랜드 마크 · 워드마크 */
/* 아래 Icon 세트(선 스타일)와 성격이 달라 분리했다. 로고는 채움(fill) 기하다.
   원본 PNG(660x660)에서 각 레이어의 바운딩박스를 재서 좌표를 역산했고,
   렌더 결과를 원본과 대조해 1~2px 이내로 맞췄다.

   기하: 가로:세로 = 1:0.546 인 둥근 마름모 하나를 세로로 균일하게 3단 배치.
   45도 회전한 정사각형에 rx 를 주면 모서리가 둥근 마름모가 되는데, 둥근
   모서리가 회전 후 꼭짓점을 안쪽으로 당기므로 정사각형 크기를 보정했다
   (유효반지름 = (H - rx)*sqrt(2) + rx).

   레이어 사이 간격은 흰 사각형이 아니라 mask 로 파낸다. 흰색으로 두면
   다크 배경에서 흰 줄이 드러난다. mask 를 쓰면 간격이 투명해진다. */
const BrandMark = ({ size = 20, title }) => {
  // 같은 화면에 마크가 둘 이상 있어도 mask id 가 겹치지 않게 한다
  const uid = useId().replace(/:/g, "");
  const cell = (cy) => (
    <g transform={`translate(24.22 ${cy}) scale(1 0.546) rotate(45)`}>
      <rect x="-15.03" y="-15.03" width="30.05" height="30.05" rx="4.51" />
    </g>
  );
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" fill="none"
      role={title ? "img" : undefined} aria-hidden={title ? undefined : true}
      style={{ display: "block", flexShrink: 0 }}>
      {title && <title>{title}</title>}
      <defs>
        <mask id={`bm-a-${uid}`} maskUnits="userSpaceOnUse" x="0" y="0" width="48" height="48">
          <rect width="48" height="48" fill="#fff" /><g fill="#000">{cell(18.18)}</g>
        </mask>
        <mask id={`bm-b-${uid}`} maskUnits="userSpaceOnUse" x="0" y="0" width="48" height="48">
          <rect width="48" height="48" fill="#fff" /><g fill="#000">{cell(27.76)}</g>
        </mask>
      </defs>
      <g fill="var(--brand-navy)"  mask={`url(#bm-b-${uid})`}>{cell(34.73)}</g>
      <g fill="var(--brand-blue)"  mask={`url(#bm-a-${uid})`}>{cell(25.15)}</g>
      <g fill="var(--brand-green)">{cell(15.56)}</g>
    </svg>
  );
};

/* 워드마크를 이미지가 아니라 실제 텍스트로 재현한다.
   확대·검색·스크린리더에 그대로 대응되고, 24px 높이에 수백 KB PNG 를
   넣지 않아도 된다. 색은 로고 원본에서 추출한 값. */
const BrandWordmark = ({ size = "var(--fs-lg)" }) => (
  <span style={{ fontSize: size, fontWeight: 800, letterSpacing: "-.02em", whiteSpace: "nowrap" }}>
    <span style={{ color: "var(--brand-ink)" }}>Career</span>
    <span style={{ color: "var(--brand-leaf)" }}> OS</span>
  </span>
);

/* ============================================================ 아이콘 (선 스타일, 와이어프레임 톤) */
const Icon = ({ name, size = 22, color = "currentColor" }) => {
  const s = { stroke: color, strokeWidth: 1.6, fill: "none", strokeLinecap: "round", strokeLinejoin: "round" };
  const paths = {
    upload: <><path d="M12 15V4" style={s} /><path d="M7 8l5-5 5 5" style={s} /><path d="M4 15v3a2 2 0 002 2h12a2 2 0 002-2v-3" style={s} /></>,
    layers: <><path d="M12 3l9 5-9 5-9-5 9-5z" style={s} /><path d="M3 13l9 5 9-5" style={s} /></>,
    archive: <><rect x="3" y="5" width="18" height="4" rx="1" style={s} /><path d="M5 9v9a2 2 0 002 2h10a2 2 0 002-2V9" style={s} /><path d="M10 13h4" style={s} /></>,
    star: <path d="M12 3l2.6 5.9 6.4.6-4.8 4.3 1.4 6.3L12 16.9l-5.6 3.2 1.4-6.3-4.8-4.3 6.4-.6L12 3z" style={s} />,
    briefcase: <><rect x="3" y="8" width="18" height="12" rx="2" style={s} /><path d="M8 8V6a2 2 0 012-2h4a2 2 0 012 2v2" style={s} /><path d="M3 13h18" style={s} /></>,
    doc: <><path d="M7 3h7l4 4v14a1 1 0 01-1 1H7a1 1 0 01-1-1V4a1 1 0 011-1z" style={s} /><path d="M14 3v4h4" style={s} /><path d="M9 12h6M9 15h6M9 9h2" style={s} /></>,
    check: <><circle cx="12" cy="12" r="9" style={s} /><path d="M8 12l3 3 5-6" style={s} /></>,
    arrowRight: <path d="M4 12h15M13 6l6 6-6 6" style={s} />,
    edit: <><path d="M4 20l1-4 11-11 3 3-11 11-4 1z" style={s} /><path d="M13 6l3 3" style={s} /></>,
    sparkle: <><path d="M12 4l1.4 4.6L18 10l-4.6 1.4L12 16l-1.4-4.6L6 10l4.6-1.4L12 4z" style={s} /></>,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24">{paths[name]}</svg>;
};

/* ============================================================ 사용 가이드 */
/* ============================================================ 타임라인 ============================================================ */
function ymToIndex(ym) { // "2024-08" -> 2024*12+8 (오래될수록 작은 수)
  if (!ym) return null;
  const [y, m] = ym.split("-").map(Number);
  if (!y) return null;
  return y * 12 + (m || 1);
}
function indexToYM(idx) {
  const y = Math.floor((idx - 1) / 12);
  const m = idx - y * 12;
  return { y, m };
}
function nowYM() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

// 겹치지 않으면 최대한 왼쪽(레인 0)에, 겹치면 다음 레인으로 배치하는 그리디 알고리즘
function assignLanes(items) {
  const sorted = [...items].sort((a, b) => a.startIdx - b.startIdx);
  const laneEnds = [];
  for (const item of sorted) {
    let placed = false;
    for (let i = 0; i < laneEnds.length; i++) {
      if (item.startIdx > laneEnds[i]) {
        item.lane = i; laneEnds[i] = item.endIdx; placed = true; break;
      }
    }
    if (!placed) { item.lane = laneEnds.length; laneEnds.push(item.endIdx); }
  }
  return sorted;
}

function makeDraftExperience(title, ym, endYm) {
  const id = "e_" + Date.now() + Math.random().toString(36).slice(2, 5);
  return {
    id, title: title || "(제목 없음)", organization: "", experienceType: "other",
    startDate: ym, endDate: endYm || ym, status: "draft", depthDone: false, usageCount: 0,
    updatedAt: new Date().toISOString().slice(0, 10), primaryCategory: "",
    competencies: [], tags: [], actions: [], context: "", assignedTask: "", discoveredProblem: "", goal: "", personalContribution: "",
    contributionLevel: "", contributionEvidence: "", coreMessage: "", oneLineSummary: "",
    completion: Object.fromEntries([...CORE_STEPS, ...DEPTH_STEPS].map(s => [s, "미입력"])),
  };
}

const TIMELINE_ROW_H = 30;
const TIMELINE_START_YM = "2021-01";

function Timeline({ experiences, setExperiences, activities, setActivities, addTrash, onOpenExp, onAnalyze, onGoArchive }) {
  const [title, setTitle] = useState("");
  const [date, setDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [filter, setFilter] = useState("all"); // all | unorganized
  const [selected, setSelected] = useState(new Set());
  // 가로 스크롤 영역이 "실제로" 넘치는지. 칸 수만 보면 데스크톱처럼 폭이 넉넉한
  // 경우에도 넘치지 않는데 안내와 탭 정지점이 생긴다.
  const laneBoxRef = useRef(null);
  const [laneOverflow, setLaneOverflow] = useState(false);

  const startIdx = ymToIndex(TIMELINE_START_YM);
  const endIdx = ymToIndex(nowYM());
  const totalRows = endIdx - startIdx + 1;

  const addActivity = () => {
    if (!title.trim() || !date) return;
    const finalEnd = endDate && endDate >= date ? endDate : date;
    setActivities(prev => [...prev, { id: "act_" + Date.now(), title: title.trim(), date, endDate: finalEnd, organized: false, linkedExpId: null }]);
    setTitle(""); setDate(""); setEndDate("");
  };

  const toggleSelect = (key) => setSelected(prev => {
    const next = new Set(prev);
    next.has(key) ? next.delete(key) : next.add(key);
    return next;
  });

  const selectedActivities = activities.filter(a => selected.has("a_" + a.id));

  const organizeSelected = () => {
    if (selectedActivities.length === 0) return;
    if (selectedActivities.length === 1) {
      const a = selectedActivities[0];
      const draft = makeDraftExperience(a.title, a.date.slice(0, 7), (a.endDate || a.date).slice(0, 7));
      setExperiences(prev => [...prev, draft]);
      setActivities(prev => prev.map(x => x.id === a.id ? { ...x, organized: true, linkedExpId: draft.id } : x));
      setSelected(new Set());
      onAnalyze(draft.id);
      return;
    }
    const newIds = [];
    setExperiences(prev => {
      const drafts = selectedActivities.map(a => makeDraftExperience(a.title, a.date.slice(0, 7), (a.endDate || a.date).slice(0, 7)));
      drafts.forEach(d => newIds.push(d.id));
      return [...prev, ...drafts];
    });
    setActivities(prev => prev.map(x => {
      const i = selectedActivities.findIndex(a => a.id === x.id);
      return i >= 0 ? { ...x, organized: true, linkedExpId: newIds[i] } : x;
    }));
    setSelected(new Set());
    onGoArchive();
  };

  const deleteSelected = () => {
    if (selectedActivities.length === 0) return;
    if (!window.confirm(`선택한 활동 ${selectedActivities.length}개를 삭제할까요?`)) return;
    selectedActivities.forEach(a => addTrash("timeline_activity", a.title, a));
    const idsToRemove = new Set(selectedActivities.map(a => a.id));
    setActivities(prev => prev.filter(a => !idsToRemove.has(a.id)));
    setSelected(new Set());
  };

  // 블록 드래그로 시기 이동 (기간 길이는 유지한 채 통째로 이동)
  const dragRef = useRef({ moved: false, justDragged: false });
  const startDrag = (it) => (e) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    const startY = e.clientY;
    const durationMonths = it.endIdx - it.startIdx;
    dragRef.current = { moved: false, justDragged: false, deltaRows: 0 };

    const onMove = (ev) => {
      const deltaRows = Math.round((ev.clientY - startY) / TIMELINE_ROW_H);
      if (deltaRows !== dragRef.current.deltaRows) dragRef.current.moved = true;
      dragRef.current.deltaRows = deltaRows;
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      const { moved, deltaRows } = dragRef.current;
      if (moved && deltaRows) {
        dragRef.current.justDragged = true;
        const newStartIdx = it.startIdx - deltaRows;
        const newEndIdx = newStartIdx + durationMonths;
        const s = indexToYM(newStartIdx), en = indexToYM(newEndIdx);
        const sStr = `${s.y}-${String(s.m).padStart(2, "0")}-01`;
        const enStr = `${en.y}-${String(en.m).padStart(2, "0")}-01`;
        if (it.kind === "activity") {
          setActivities(prev => prev.map(a => a.id === it.raw.id ? { ...a, date: sStr, endDate: enStr } : a));
        } else {
          setExperiences(prev => prev.map(x => x.id === it.raw.id ? { ...x, startDate: sStr.slice(0, 7), endDate: enStr.slice(0, 7) } : x));
        }
      }
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  };
  const handleBlockClick = (it) => {
    if (dragRef.current.justDragged) { dragRef.current.justDragged = false; return; }
    toggleSelect(it.key);
  };

  // 경험 + (미정리 필터가 아니면 정리된 활동도 숨김 — 이미 경험으로 존재하므로 중복 방지)
  const expItems = experiences.filter(e => e.startDate).map(e => {
    const s = ymToIndex(e.startDate.slice(0, 7));
    const en = ymToIndex((e.endDate || e.startDate).slice(0, 7)) || s;
    return { key: "e_" + e.id, kind: "experience", title: e.title, startIdx: Math.max(s, startIdx), endIdx: Math.min(Math.max(en, s), endIdx), raw: e };
  }).filter(it => it.startIdx <= endIdx && it.endIdx >= startIdx);

  const actItems = activities.filter(a => !a.organized).map(a => {
    const s = ymToIndex(a.date.slice(0, 7));
    const en = Math.max(ymToIndex((a.endDate || a.date).slice(0, 7)) || s, s);
    return { key: "a_" + a.id, kind: "activity", title: a.title, startIdx: Math.max(s, startIdx), endIdx: Math.min(en, endIdx), raw: a };
  }).filter(it => it.startIdx <= endIdx);

  const allItems = filter === "unorganized" ? actItems : [...expItems, ...actItems];
  const lanedItems = assignLanes(allItems);
  const laneCount = Math.max(1, ...lanedItems.map(it => it.lane + 1));

  useEffect(() => {
    const el = laneBoxRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const check = () => setLaneOverflow(el.scrollWidth > el.clientWidth + 1);
    check();
    const ro = new ResizeObserver(check);
    ro.observe(el);
    return () => ro.disconnect();
  }, [laneCount]);

  const rows = [];
  for (let idx = endIdx; idx >= startIdx; idx--) rows.push(idx);

  return (
    <div>
      <H2>타임라인</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.6 }}>
        연도·월을 쭉 훑어보면서, 아직 경험 보관함에 정리하지 않은 활동을 빠르게 기록하고 골라서 정리하세요.
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }} className="wrap-sm">
          <Input placeholder="이때 무슨 일이 있었나요? (예: 팀 프로젝트 발표)" value={title} onChange={e => setTitle(e.target.value)} style={{ flex: 1, minWidth: 200 }} />
          {/* 시작~종료는 한 쌍이라 좁은 화면에서도 같은 줄에 둔다.
              고정 140px 이면 두 개가 들어가지 않아 세로로 쪼개졌다.
              남는 폭을 반씩 나눠 갖도록 flex 로 바꾸고 wrap 을 뺐다. */}
          <div style={{ display: "flex", alignItems: "center", gap: 6, flex: "1 1 260px", minWidth: 0 }}>
            <input type="month" value={date ? date.slice(0, 7) : ""} onChange={e => setDate(e.target.value ? e.target.value + "-01" : "")}
              style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, flex: 1, minWidth: 0 }} />
            <span style={{ fontSize: "var(--fs-sm)", color: C.faintText, flexShrink: 0 }}>~</span>
            <input type="month" value={endDate ? endDate.slice(0, 7) : ""} min={date ? date.slice(0, 7) : undefined}
              onChange={e => setEndDate(e.target.value ? e.target.value + "-01" : "")}
              style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, flex: 1, minWidth: 0 }} />
          </div>
          <Btn primary disabled={!title.trim() || !date} onClick={addActivity}>추가</Btn>
        </div>
        <div style={{ fontSize: "var(--fs-xs)", color: C.faintText, marginTop: 6 }}>종료 년월은 선택 사항입니다 — 비워두면 하루·한 달짜리 활동(점)으로, 채우면 기간이 있는 활동(막대)으로 표시됩니다.</div>
      </Card>

      <div style={{ display: "flex", gap: 6, marginBottom: 16 }}>
        {[["all", "전체"], ["unorganized", "미정리만"]].map(([v, l]) => (
          <button key={v} onClick={() => setFilter(v)} style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-5)", borderRadius: "var(--r-lg)", cursor: "pointer",
            border: `1px solid ${filter === v ? C.text : C.line}`, background: filter === v ? C.text : C.panel, color: filter === v ? "#fff" : C.sub }}>{l}</button>
        ))}
      </div>

      <div style={{ display: "flex", gap: 16, marginBottom: 6 }}>
        <span style={{ fontSize: "var(--fs-xs)", color: C.sub }}><span style={{ display: "inline-block", width: 8, height: 8, borderRadius: "var(--r-full)", background: C.panel, border: `2px solid ${C.sub}`, marginRight: 5 }} />미정리 (점선/테두리만)</span>
        <span style={{ fontSize: "var(--fs-xs)", color: C.sub }}><span style={{ display: "inline-block", width: 12, height: 8, borderRadius: "var(--r-xs)", background: C.greenBg, border: `1px solid ${C.green}`, marginRight: 5 }} />정리된 경험 (채움)</span>
      </div>

      <div style={{ display: "flex" }}>
        <div style={{ width: 52, flexShrink: 0 }}>
          {rows.map((idx, i) => {
            const { y, m } = indexToYM(idx);
            const isJan = m === 1;
            const isTop = i === 0;
            return (
              <div key={idx} style={{ height: TIMELINE_ROW_H, display: "flex", alignItems: "center", fontSize: "var(--fs-xs)", color: C.faintText,
                borderTop: i === 0 ? "none" : `1px solid ${C.lineSoft}` }}>
                {/* "2026·1월" 을 한 줄에 쓰면 52px 컬럼을 넘겨 "2026·1 / 월" 로 갈라진다.
                    연도를 굵게 위에, 월은 다른 달과 같은 서체로 아래에 둔다. */}
                {(isJan || isTop) ? (
                  <span style={{ display: "flex", flexDirection: "column", lineHeight: 1.1 }}>
                    <span style={{ fontWeight: 700, color: C.text, fontSize: "var(--fs-2xs)" }}>{y}</span>
                    <span>{m}월</span>
                  </span>
                ) : `${m}월`}
              </div>
            );
          })}
        </div>

        <div style={{ flex: 1, minWidth: 0 }}>
          {laneCount > 1 && (
            <div style={{ marginBottom: 6 }}>
              <span style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>
                동시에 진행된 활동이 {laneCount}칸으로 나뉘어 있습니다{laneOverflow ? " — 옆으로 밀어서 보세요" : ""}
              </span>
            </div>
          )}
          {/* 가로 스크롤 영역. 포커스를 받을 수 있어야 키보드로도 좌우 이동이 된다
              (브라우저가 포커스된 overflow 컨테이너를 화살표 키로 스크롤한다).
              예전에는 ◀▶ 버튼이 유일한 키보드 경로였는데, 터치·마우스에서는
              스와이프로 충분해 버튼을 걷어내고 이 방식으로 바꿨다.
              넘치지 않을 땐 스크롤할 것이 없으므로 탭 순서에 넣지 않는다. */}
          <div
            ref={laneBoxRef}
            {...(laneOverflow ? { tabIndex: 0, role: "region", "aria-label": `타임라인 활동 ${laneCount}칸, 좌우로 스크롤` } : {})}
            style={{ overflowX: "auto", scrollbarWidth: "thin" }}>
          <div style={{ position: "relative", height: totalRows * TIMELINE_ROW_H, display: "flex", gap: 12, paddingLeft: 12, borderLeft: `1px solid ${C.line}`, minWidth: laneCount * 162 }}>
            {Array.from({ length: laneCount }).map((_, laneIdx) => (
              <div key={laneIdx} style={{ position: "relative", width: 150, flexShrink: 0 }}>
                {rows.map((idx, i) => (
                  <div key={idx} style={{ position: "absolute", top: i * TIMELINE_ROW_H, left: 0, right: 0, height: 1, background: i === 0 ? "transparent" : C.lineSoft }} />
                ))}
                {lanedItems.filter(it => it.lane === laneIdx).map(it => {
                  const topRow = endIdx - it.endIdx;
                  const bottomRow = endIdx - it.startIdx;
                  const top = topRow * TIMELINE_ROW_H + 3;
                  const height = (bottomRow - topRow + 1) * TIMELINE_ROW_H - 6;
                  const isDot = it.startIdx === it.endIdx && it.kind === "activity";
                  const isSelected = selected.has(it.key);
                  const isOrganized = it.kind === "experience";
                  if (isDot) {
                    return (
                      <div key={it.key} onMouseDown={startDrag(it)} {...clickableProps(() => handleBlockClick(it))} title={it.title + " (드래그해서 시기 이동)"}
                        style={{ position: "absolute", top: top + 6, left: 2, right: 2, display: "flex", alignItems: "flex-start", gap: 7, cursor: "grab" }}>
                        <span style={{ width: 10, height: 10, borderRadius: "var(--r-full)", background: isSelected ? C.text : C.panel, border: `2px solid ${isSelected ? C.text : C.sub}`, flexShrink: 0, marginTop: 2 }} />
                        <span style={{ fontSize: "var(--fs-sm)", fontWeight: isSelected ? 700 : 500, color: C.text, lineHeight: 1.35, display: "-webkit-box", WebkitLineClamp: 3, WebkitBoxOrient: "vertical", overflow: "hidden" }}>{it.title}</span>
                      </div>
                    );
                  }
                  return (
                    <div key={it.key} onMouseDown={startDrag(it)} {...clickableProps(() => handleBlockClick(it))} title={it.title + " (드래그해서 시기 이동)"} style={{
                      position: "absolute", top, left: 3, right: 3, height: Math.max(height, 24), borderRadius: "var(--r-sm)", cursor: "grab", boxSizing: "border-box",
                      background: isOrganized ? C.greenBg : C.panel,
                      border: isOrganized ? `1px solid ${C.green}` : `1.5px dashed ${isSelected ? C.text : C.sub}`,
                      outline: isSelected ? `2px solid ${C.text}` : "none", outlineOffset: 1,
                      padding: "var(--sp-2) var(--sp-3)", fontSize: "var(--fs-sm)", fontWeight: isSelected ? 700 : 500, color: isOrganized ? C.greenText : C.text, lineHeight: 1.35,
                      display: "-webkit-box", WebkitLineClamp: Math.max(1, Math.floor((Math.max(height, 24) - 12) / 16)), WebkitBoxOrient: "vertical", overflow: "hidden" }}>
                      {it.title}
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
          </div>
        </div>
      </div>

      {selected.size === 1 && (() => {
        const key = [...selected][0];
        if (key.startsWith("a_")) {
          const a = activities.find(x => "a_" + x.id === key);
          if (!a) return null;
          return <TimelineActivityPanel activity={a} setActivities={setActivities} onDeselect={() => setSelected(new Set())}
            onOrganize={() => organizeSelected()} onDelete={() => deleteSelected()} />;
        }
        const e = experiences.find(x => "e_" + x.id === key);
        if (!e) return null;
        return <TimelineExperienceNote exp={e} setExperiences={setExperiences} onOpenExp={onOpenExp} onDeselect={() => setSelected(new Set())} />;
      })()}

      {selected.size > 1 && (
        <div style={{ position: "sticky", bottom: 16, marginTop: 16, background: C.panel, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)", padding: "var(--sp-4) var(--sp-5)", display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 2px 10px rgba(0,0,0,0.06)" }} className="wrap-sm">
          <span style={{ fontSize: "var(--fs-sm)" }}>{selected.size}개 선택됨{selectedActivities.length < selected.size ? " (정리된 경험은 일괄 작업 대상에서 제외)" : ""}</span>
          <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
            <Btn small onClick={() => setSelected(new Set())}>선택 해제</Btn>
            <Btn small onClick={deleteSelected} disabled={selectedActivities.length === 0}>삭제</Btn>
            <Btn small primary onClick={organizeSelected} disabled={selectedActivities.length === 0}>선택한 항목 정리하기</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

function TimelineActivityPanel({ activity, setActivities, onDeselect, onOrganize, onDelete }) {
  const [title, setTitle] = useState(activity.title);
  const [date, setDate] = useState(activity.date.slice(0, 7));
  const [endDate, setEndDate] = useState((activity.endDate || activity.date).slice(0, 7));

  const save = () => {
    if (!title.trim() || !date) return;
    const finalEnd = endDate && endDate >= date ? endDate : date;
    setActivities(prev => prev.map(a => a.id === activity.id ? { ...a, title: title.trim(), date: date + "-01", endDate: finalEnd + "-01" } : a));
  };

  return (
    <Card style={{ marginTop: 16, background: C.accent }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <Label>미정리 활동 — 수정</Label>
        <span {...clickableProps(onDeselect, { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
      </div>
      <Input value={title} onChange={e => setTitle(e.target.value)} style={{ marginBottom: 8 }} />
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }} className="wrap-sm">
        <input type="month" value={date} onChange={e => setDate(e.target.value)}
          style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, width: 140 }} />
        <span style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>~</span>
        <input type="month" value={endDate} min={date} onChange={e => setEndDate(e.target.value)}
          style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, width: 140 }} />
      </div>
      <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
        <Btn small primary onClick={save}>저장</Btn>
        <Btn small onClick={onOrganize}>정리하기 →</Btn>
        <Btn small onClick={onDelete}>삭제</Btn>
      </div>
    </Card>
  );
}

function TimelineExperienceNote({ exp, setExperiences, onOpenExp, onDeselect }) {
  const [note, setNote] = useState(exp.rawNote || "");
  const [title, setTitle] = useState(exp.title);
  const [startYm, setStartYm] = useState((exp.startDate || "").slice(0, 7));
  const [endYm, setEndYm] = useState((exp.endDate || exp.startDate || "").slice(0, 7));
  const autosave = useAutosave(note);

  useEffect(() => {
    const t = setTimeout(() => {
      setExperiences(prev => prev.map(e => e.id === exp.id ? { ...e, rawNote: note } : e));
    }, 500);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [note]);

  const saveDates = () => {
    if (!title.trim() || !startYm) return;
    const finalEnd = endYm && endYm >= startYm ? endYm : startYm;
    setExperiences(prev => prev.map(e => e.id === exp.id ? { ...e, title: title.trim(), startDate: startYm, endDate: finalEnd } : e));
  };

  return (
    <Card style={{ marginTop: 16, background: C.greenBg }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <Label>정리된 경험 — 제목·시기 수정</Label>
        <span {...clickableProps(onDeselect, { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
      </div>
      <Input value={title} onChange={e => setTitle(e.target.value)} style={{ marginBottom: 8, fontWeight: 700 }} />
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }} className="wrap-sm">
        <input type="month" value={startYm} onChange={e => setStartYm(e.target.value)}
          style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, width: 140 }} />
        <span style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>~</span>
        <input type="month" value={endYm} min={startYm} onChange={e => setEndYm(e.target.value)}
          style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, width: 140 }} />
        <Btn small primary onClick={saveDates}>저장</Btn>
      </div>
      <Label>관련 메모</Label>
      <Textarea value={note} onChange={e => setNote(e.target.value)} placeholder="이 경험과 관련해서 떠오른 걸 자유롭게 적어두세요" rows={4} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
        <AutosaveIndicator state={autosave} />
        <Btn small onClick={() => onOpenExp(exp.id)}>전체 경험 보기 →</Btn>
      </div>
    </Card>
  );
}
function PersonalAssistant({ experiences, skills, certs, awards, resumeProfile, applications, metrics, history, setHistory, onGo }) {
  return (
    <div style={{ maxWidth: 720 }}>
      <H2>AI에게 물어보기</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.6 }}>
        취업 준비하면서 드는 사소한 질문이나 개인적인 고민을 편하게 물어보세요. 정리해두신 경험·역량·지원 현황을 참고해서 답합니다. 채용담당자처럼 평가하는 곳이 아니라, 옆에서 같이 생각해보는 곳입니다.
      </div>
      <EssayChat
        title="AI에게 물어보기"
        subtitle="내 정보를 참고해서 답합니다"
        systemPrompt={PERSONAL_ASSISTANT_SYSTEM_PROMPT}
        contextText={buildPersonalContext(experiences, skills, certs, awards, resumeProfile, applications, metrics)}
        autoStartMessage="안녕! 요즘 취업 준비하면서 궁금한 거나 고민되는 거 있으면 편하게 물어봐."
        inputPlaceholder="예: 내 경험 중에 뭐가 제일 강점인 것 같아? / 이 회사 지원할까 말까 고민돼"
        onClose={() => onGo("home")}
        closeLabel="← 홈으로"
        history={history}
        onHistoryChange={setHistory}
      />
    </div>
  );
}

function Guide({ onGo }) {
  const flow = [
    { icon: "upload", title: "자료 준비", desc: "기존 이력서·메모 파일을 가져오거나, 경험을 새로 등록" },
    { icon: "layers", title: "경험 분석", desc: "핵심 5단계 질문에 답하며 사실을 구조화" },
    { icon: "star", title: "역량·스킬 연결", desc: "도구·역량에 경험 근거를 연결" },
    { icon: "briefcase", title: "지원 등록", desc: "회사·직무별로 요구 역량과 경험을 매칭" },
    { icon: "doc", title: "자소서·면접 준비", desc: "문항별 문장 작성, 예상 질문 연습" },
    { icon: "check", title: "최종 이력서", desc: "승인된 문장만 모아 완성" },
  ];

  const tabs = [
    { icon: "upload", nav: "import", title: "파일 가져오기", desc: "기존 이력서·정리 파일(.docx/.xlsx/.txt)에서 AI가 초안을 추출합니다. 모든 항목은 반영 전 직접 확인·수정합니다." },
    { icon: "layers", nav: "analyze", title: "경험 분석", desc: "배경·문제·행동·기여도·성과 5단계로 경험을 구조화합니다. 심화 4단계는 나중에 채워도 됩니다." },
    { icon: "archive", nav: "archive", title: "경험 보관함", desc: "등록한 모든 경험을 경험별·역량별·질문별로 찾아봅니다." },
    { icon: "star", nav: "skills", title: "역량·스킬", desc: "도구·역량마다 실제로 할 수 있는 일을 적고, 근거가 되는 경험을 연결합니다." },
    { icon: "briefcase", nav: "apply", title: "지원 관리", desc: "지원할 회사마다 요구 역량 매칭, 자소서 문항, 면접 질문을 따로 관리합니다." },
    { icon: "doc", nav: "resume", title: "기본 이력서", desc: "경험 보관함에서 승인된 문장과 근거가 연결된 역량이 자동으로 모이고, 자격증·어학·수상기록은 여기서 직접 관리합니다." },
  ];

  const approvalFlow = [
    { label: "AI 초안", color: C.ai, desc: "AI가 문장을 생성한 직후" },
    { label: "수정 중", color: C.blueText, desc: "직접 내용을 고치는 단계" },
    { label: "승인됨", color: C.greenText, desc: "확인 완료 — 이력서에 사용 가능" },
  ];

  return (
    <div style={{ maxWidth: 880 }}>
      <H2>사용 가이드</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 24, lineHeight: 1.6 }}>
        처음 쓰는 분들을 위한 전체 흐름입니다. 순서대로 하지 않아도 되지만, 처음이라면 이 순서를 추천합니다.
      </div>

      {/* 전체 흐름 */}
      <Card style={{ marginBottom: 20 }}>
        <Label>전체 흐름</Label>
        <div style={{ display: "flex", alignItems: "stretch", gap: 4, marginTop: 10, flexWrap: "wrap" }}>
          {flow.map((f, i) => (
            <React.Fragment key={f.title}>
              <div style={{ flex: "1 1 140px", minWidth: 130, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)", padding: "var(--sp-5) var(--sp-5)", textAlign: "center", background: C.bg }}>
                <div style={{ display: "flex", justifyContent: "center", marginBottom: 8, color: C.blueText }}>
                  <Icon name={f.icon} size={26} />
                </div>
                <div style={{ fontSize: "var(--fs-base)", fontWeight: 700, marginBottom: 4 }}>{i + 1}. {f.title}</div>
                <div style={{ fontSize: "var(--fs-xs)", color: C.sub, lineHeight: 1.5 }}>{f.desc}</div>
              </div>
              {i < flow.length - 1 && (
                <div style={{ display: "flex", alignItems: "center", color: C.faintText, flex: "0 0 auto" }}>
                  <Icon name="arrowRight" size={18} />
                </div>
              )}
            </React.Fragment>
          ))}
        </div>
      </Card>

      {/* 탭별 설명 */}
      <Label>탭별 안내</Label>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 8, marginBottom: 20 }} className="stack-sm">
        {tabs.map(t => (
          <Card key={t.nav} onClick={() => onGo(t.nav)} style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
            <div style={{ color: C.blueText, flexShrink: 0, marginTop: 2 }}><Icon name={t.icon} size={24} /></div>
            <div>
              <div style={{ fontSize: "var(--fs-md)", fontWeight: 700, marginBottom: 4 }}>{t.title}</div>
              <div style={{ fontSize: "var(--fs-sm)", color: C.sub, lineHeight: 1.55 }}>{t.desc}</div>
            </div>
          </Card>
        ))}
      </div>

      {/* 승인 개념 */}
      <Card>
        <Label>"승인"이 왜 필요한가요?</Label>
        <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 14, lineHeight: 1.6 }}>
          AI가 쓴 문장을 그대로 이력서에 쓰지 않도록, 사실(경험 기록)과 표현(AI가 다듬은 문장)을 분리했습니다.
          문장은 아래 3단계를 거쳐야 최종 이력서·자소서에 쓸 수 있습니다.
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap" }}>
          {approvalFlow.map((s, i) => (
            <React.Fragment key={s.label}>
              <div style={{ flex: "1 1 160px", border: `1px solid ${C.line}`, borderTop: `3px solid ${s.color}`, borderRadius: "var(--r-lg)", padding: "var(--sp-5) var(--sp-5)" }}>
                <div style={{ fontSize: "var(--fs-base)", fontWeight: 700, color: s.color, marginBottom: 4 }}>{s.label}</div>
                <div style={{ fontSize: "var(--fs-xs)", color: C.sub, lineHeight: 1.5 }}>{s.desc}</div>
              </div>
              {i < approvalFlow.length - 1 && <div style={{ color: C.faintText }}><Icon name="arrowRight" size={16} /></div>}
            </React.Fragment>
          ))}
        </div>
      </Card>

      {/* 데이터 저장/전송 안내 */}
      <Card style={{ marginTop: 20 }}>
        <Label>내 데이터는 어디에 저장되고, 어디로 전송되나요</Label>
        <div style={{ fontSize: "var(--fs-base)", color: C.sub, lineHeight: 1.8 }}>
          · 모든 데이터는 <b>이 브라우저에만</b> 저장됩니다 (서버 저장 없음). 다른 기기·다른 브라우저·시크릿 모드에서는 보이지 않습니다.<br />
          · 브라우저 데이터를 지우면 복구할 수 없습니다. 사이드바 하단의 <b>"데이터 백업"</b>을 주기적으로 받아두는 걸 권장합니다.<br />
          · <b>"파일 가져오기"</b>와 <b>"자소서·면접 챗봇"</b>을 사용하면, 그 순간 입력한 내용이 AI 응답 생성을 위해 외부 AI 서버(Anthropic/OpenAI/Gemini 중 설정된 곳)로 전송됩니다. 전송된 내용은 응답 생성에만 쓰이고 이 앱이 별도로 저장하지 않습니다.<br />
          · 백업 파일은 암호화되지 않은 평문 JSON입니다. 공유 컴퓨터에 저장하지 않도록 주의하세요.
        </div>
      </Card>
    </div>
  );
}

/* ============================================================ 첫 방문자 온보딩 화면 */
function HomeOnboarding({ onGoAnalyze, onGoImport, onLoadDemo, onGoGuide }) {
  return (
    <div style={{ maxWidth: 560, margin: "40px auto 0" }}>
      <div style={{ textAlign: "center", marginBottom: 28 }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 14 }}>
          <BrandMark size={40} />
        </div>
        <h1 style={{ fontSize: "var(--fs-3xl)", fontWeight: 800, margin: "0 0 8px" }}>Career OS에 오신 걸 환영합니다</h1>
        <div style={{ fontSize: "var(--fs-base)", color: C.sub, lineHeight: 1.6 }}>
          경험을 정리하고, 이력서·자소서로 이어가는 걸 도와드립니다.
        </div>
      </div>

      <Card onClick={onGoImport} style={{
        textAlign: "center", padding: "var(--sp-8) var(--sp-7)", marginBottom: 14,
        border: `1px solid ${C.primary}`, background: C.primaryBg }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 12, color: C.primary }}>
          <Icon name="upload" size={30} />
        </div>
        <div style={{ fontSize: "var(--fs-lg)", fontWeight: 700, marginBottom: 6 }}>시작하기</div>
        <div style={{ fontSize: "var(--fs-sm)", color: C.sub, lineHeight: 1.5 }}>
          이미 작성된 자소서 등의 파일이 있다면 첨부해주세요! 저희가 정리해드립니다.<br />
          없다면 그냥 아무거나 적어도 괜찮아요 — 빈 화면부터 시작할 수도 있습니다.
        </div>
      </Card>

      <div style={{ display: "flex", justifyContent: "center", gap: 18, fontSize: "var(--fs-sm)" }}>
        <span {...clickableProps(onGoGuide)} style={{ color: C.sub, textDecoration: "underline", cursor: "pointer" }}>전체 사용법 먼저 보기</span>
        <span {...clickableProps(onLoadDemo)} style={{ color: C.sub, textDecoration: "underline", cursor: "pointer" }}>예시 데이터로 먼저 둘러보기</span>
      </div>
    </div>
  );
}

function Home({ experiences, applications, onGoAnalyze, onGoImport, onOpenDetail, onOpenApp, isBlankSlate, onLoadDemo, onGoGuide }) {
  if (isBlankSlate) {
    return <HomeOnboarding onGoAnalyze={onGoAnalyze} onGoImport={onGoImport} onLoadDemo={onLoadDemo} onGoGuide={onGoGuide} />;
  }

  const total = experiences.length;
  const done = experiences.filter(e => e.status === "complete").length;
  const needs = experiences.filter(e => e.status === "needs_revision").length;
  const draft = experiences.filter(e => e.status === "draft").length;

  const typeCoverage = [
    ["데이터 분석 경험", experiences.filter(e => e.competencies.includes("데이터 분석") || e.competencies.includes("데이터 관리")).length],
    ["리더십 경험", experiences.filter(e => e.competencies.includes("리더십")).length],
    ["어려움 극복 경험", experiences.filter(e => !!e.difficulty).length],
    ["협업 경험", experiences.filter(e => e.actions?.some(a => a.actionType === "collaboration")).length],
  ];

  const nextActions = (() => {
    const acts = [];
    experiences.filter(e => e.status === "needs_revision").forEach(e =>
      acts.push({ text: `「${e.title}」 부족한 부분 보완하기`, act: () => onOpenDetail(e.id) }));
    experiences.filter(e => e.status === "draft").forEach(e =>
      acts.push({ text: `「${e.title}」 경험 분석 시작하기`, act: () => onOpenDetail(e.id) }));
    experiences.filter(e => e.status === "complete" && !e.depthDone).forEach(e =>
      acts.push({ text: `「${e.title}」 심화 단계(어려움·배운 점) 입력하기`, act: () => onOpenDetail(e.id) }));
    applications.forEach(a => {
      if ((a.interviews || []).some(iq => !iq.selectedExperienceId)) {
        acts.push({ text: `${a.company} 면접 질문에 사용할 경험 선택하기`, act: () => onOpenApp(a.id) });
      }
      if ((a.essays || []).some(q => q.status === "not_started")) {
        acts.push({ text: `${a.company} 자소서 문항 작성 시작하기`, act: () => onOpenApp(a.id) });
      }
    });
    return acts.slice(0, 5);
  })();

  return (
    <div>
      <h1 style={{ fontSize: "var(--fs-3xl)", fontWeight: 800, margin: "0 0 4px" }}>오늘 할 일부터 시작하세요</h1>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 22 }}>차트보다 행동. 다음에 해야 할 일을 바로 보여드립니다.</div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }} className="stack-sm">
        {/* 다음 행동 */}
        <Card style={{ gridColumn: "1 / -1", background: C.lineSoft, border: "none" }}>
          <Label>다음 행동</Label>
          {nextActions.length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.sub, padding: "var(--sp-3) 0" }}>지금 당장 처리할 일이 없습니다. 새 경험을 등록하거나 지원을 추가해보세요.</div>}
          {nextActions.map((a, i) => (
            <div key={i} {...clickableProps(a.act)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "var(--sp-4) 0", borderBottom: i < nextActions.length - 1 ? `1px solid ${C.line}` : "none", cursor: "pointer" }}>
              <span style={{ fontSize: "var(--fs-md)" }}>{a.text}</span>
              <span style={{ color: C.faintText, fontSize: "var(--fs-base)" }}>→</span>
            </div>
          ))}
        </Card>

        {/* 준비 현황 */}
        <Card>
          <Label>경험 준비 현황</Label>
          <div style={{ display: "flex", gap: 22, marginTop: 8 }}>
            {/* 이 숫자들은 24px/800 로 렌더되는 "텍스트"다. 채움용 C.green / C.orange 를
                 그대로 쓰면 각각 2.32:1 / 1.85:1 로 큰 텍스트 기준 3:1 에도 못 미친다.
                 텍스트용 변형을 쓴다. */}
            {[["전체", total, C.text], ["분석 완료", done, C.greenText], ["보완 필요", needs, C.orangeText], ["초기 메모", draft, C.sub]].map(([l, v, c]) => (
              <div key={l}>
                <div style={{ fontSize: "var(--fs-4xl)", fontWeight: 800, color: c }}>{v}</div>
                <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{ marginTop: 14, height: 6, background: C.lineSoft, borderRadius: "var(--r-xs)", overflow: "hidden", display: "flex" }}>
            <div style={{ width: `${(done / total) * 100}%`, background: C.green }} />
            <div style={{ width: `${(needs / total) * 100}%`, background: C.orange }} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }} className="wrap-sm">
            <Btn small primary onClick={onGoAnalyze}>+ 새 경험 분석 시작</Btn>
            <Btn small onClick={onGoImport}>파일에서 가져오기</Btn>
          </div>
        </Card>

        {/* 부족한 경험 유형 */}
        <Card>
          <Label>부족한 경험 유형</Label>
          {typeCoverage.map(([l, n]) => (
            <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "var(--sp-2) 0", fontSize: "var(--fs-base)" }}>
              <span>{l}</span>
              {n === 0 ? <Badge label="준비 부족" color={C.redText} bg={C.redBg} /> : <span style={{ fontWeight: 700 }}>{n}개</span>}
            </div>
          ))}
          <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginTop: 6 }}>실패·갈등 경험은 면접 단골 질문입니다.</div>
        </Card>

        {/* 진행 중 지원 */}
        <Card style={{ gridColumn: "1 / -1" }}>
          <Label>진행 중인 지원</Label>
          {applications.map(a => (
            <div key={a.id} {...clickableProps(() => onOpenApp(a.id))} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 90px 100px 160px", alignItems: "center", gap: 10, padding: "var(--sp-4) 0", borderBottom: `1px solid ${C.lineSoft}`, cursor: "pointer", fontSize: "var(--fs-base)" }} className={"ui-click tbl-row"}>
              <div style={{ fontWeight: 700 }}>{a.company} <span style={{ fontWeight: 400, color: C.sub }}>{a.position}</span></div>
              <div style={{ color: C.sub }}>마감 {a.deadline}</div>
              <Badge label={{ interested: "관심", analyzing: "분석 중", writing: "작성 중", submitted: "제출", interview: "면접", result: "결과" }[a.status]} color={C.blueText} bg={C.blueBg} />
              <div style={{ fontSize: "var(--fs-sm)", color: C.sub }}>자소서 {a.essayProgress}%</div>
              <div style={{ height: 5, background: C.lineSoft, borderRadius: "var(--r-xs)" }}><div style={{ width: `${a.essayProgress}%`, height: "100%", background: C.blue, borderRadius: "var(--r-xs)" }} /></div>
            </div>
          ))}
        </Card>
      </div>
    </div>
  );
}

/* ============================================================ 경험 분석 */
function Analyze({ experiences, setExperiences, analyzeId, setAnalyzeId, metrics, setMetrics, onDone }) {
  const exp = experiences.find(e => e.id === analyzeId);
  if (!exp) return <AnalyzeStart experiences={experiences} setExperiences={setExperiences} onStart={setAnalyzeId} />;
  return <AnalyzeFlow exp={exp} setExperiences={setExperiences} metrics={metrics} setMetrics={setMetrics} onExit={() => setAnalyzeId(null)} onDone={onDone} />;
}

function AnalyzeStart({ experiences, setExperiences, onStart }) {
  const [form, setForm] = useState({ title: "", organization: "", role: "", experienceType: "internship", rawNote: "" });
  const drafts = experiences.filter(e => e.status !== "complete");
  const canStart = form.title.trim() && form.rawNote.trim();

  const create = () => {
    const id = "e_" + Date.now();
    setExperiences(p => [...p, { id, ...form, status: "draft", depthDone: false, usageCount: 0, updatedAt: "2026-07-21", primaryCategory: "",
      competencies: [], tags: [], actions: [], context: "", assignedTask: "", discoveredProblem: "", goal: "", personalContribution: "",
      contributionLevel: "", contributionEvidence: "", coreMessage: "", oneLineSummary: "",
      completion: Object.fromEntries([...CORE_STEPS, ...DEPTH_STEPS].map(s => [s, "미입력"])) }]);
    onStart(id);
  };

  return (
    <div style={{ maxWidth: 640 }}>
      <H2>새 경험 분석</H2>
      <Card>
        <div style={{ fontSize: "var(--fs-base)", color: C.sub, lineHeight: 1.7, marginBottom: 16, padding: "var(--sp-5)", background: C.bg, borderRadius: "var(--r-lg)" }}>
          처음부터 완벽하게 작성할 필요는 없습니다.<br />기억나는 내용을 자유롭게 적으면, 질문을 통해 함께 구체화합니다.
        </div>
        <div style={{ display: "grid", gap: 13 }}>
          <div><Label>경험 제목 *</Label><Input placeholder="예: 웹사이트 운영 프로모션 기획" value={form.title} onChange={e => setForm(f => ({ ...f, title: e.target.value }))} /></div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }} className="stack-sm">
            <div><Label>소속 또는 활동명</Label><Input placeholder="예: 온라인 쇼핑몰 (인턴)" value={form.organization} onChange={e => setForm(f => ({ ...f, organization: e.target.value }))} /></div>
            <div><Label>당시 역할</Label><Input placeholder="예: E-commerce Assistant" value={form.role} onChange={e => setForm(f => ({ ...f, role: e.target.value }))} /></div>
          </div>
          <div>
            <Label>경험 유형</Label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {[["internship", "인턴"], ["full_time", "정규직"], ["part_time", "아르바이트"], ["school_project", "학교 프로젝트"], ["external_activity", "대외활동"], ["club", "동아리"], ["competition", "공모전"], ["personal_project", "개인 프로젝트"], ["other", "기타"]].map(([v, l]) => (
                <button key={v} onClick={() => setForm(f => ({ ...f, experienceType: v }))} style={{
                  fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-4)", borderRadius: "var(--r-lg)", cursor: "pointer",
                  border: `1px solid ${form.experienceType === v ? C.text : C.line}`,
                  background: form.experienceType === v ? C.text : C.panel, color: form.experienceType === v ? "#fff" : C.sub }}>
                  {l}
                </button>
              ))}
            </div>
          </div>
          <div><Label>자유 메모 *</Label>
            <Textarea rows={5} placeholder={"예: 아르바이트 근무 중 매장 프로모션을 진행했다.\n기존 할인만 하는 것보다 사은품을 주는 게 좋을 것 같았다.\n과거 데이터를 분석해서 제품을 골랐고 매출이 올랐다."}
              value={form.rawNote} onChange={e => setForm(f => ({ ...f, rawNote: e.target.value }))} />
          </div>
        </div>
        <div style={{ marginTop: 16, display: "flex", justifyContent: "flex-end" }}>
          <Btn primary disabled={!canStart} onClick={create}>단계별 분석 시작 →</Btn>
        </div>
      </Card>

      {drafts.length > 0 && (
        <div style={{ marginTop: 22 }}>
          <Label>분석 이어하기</Label>
          {drafts.map(e => (
            <Card key={e.id} onClick={() => onStart(e.id)} style={{ marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center", padding: "var(--sp-5) var(--sp-6)" }}>
              <div>
                <span style={{ fontWeight: 700, fontSize: "var(--fs-base)" }}>{e.title}</span>
                <span style={{ color: C.faintText, fontSize: "var(--fs-sm)", marginLeft: 8 }}>{e.organization}</span>
              </div>
              <Badge label={STATUS_LABEL[e.status]} color={STATUS_COLOR[e.status][0]} bg={STATUS_COLOR[e.status][1]} />
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function AnalyzeFlow({ exp, setExperiences, metrics, setMetrics, onExit, onDone }) {
  const [stepIdx, setStepIdx] = useState(0);
  const [showDepth, setShowDepth] = useState(false);
  const steps = showDepth ? DEPTH_STEPS : CORE_STEPS;
  const step = steps[stepIdx];
  const [local, setLocal] = useState({ ...exp });
  const autosave = useAutosave(JSON.stringify(local));

  const patch = (k, v) => setLocal(p => ({ ...p, [k]: v }));
  const commit = (extra = {}) => setExperiences(prev => prev.map(e => e.id === exp.id ? { ...e, ...local, ...extra, updatedAt: new Date().toISOString().slice(0, 10) } : e));

  const STEP_MIN_LENGTH = { 배경: 8, 문제: 8, 기여도: 8, 목표: 4, 어려움: 8, "배운 점": 8, "직무 연결": 4 };
  const isStepFilled = (name) => {
    if (name === "행동") return (local.actions || []).length > 0;
    if (name === "성과") return !!(local.oneLineSummary || "").trim() || !!(local.qualitative || "").trim()
      || metrics.some(m => m.experienceId === exp.id);
    const fieldMap = { 배경: "context", 문제: "discoveredProblem", 기여도: "contributionEvidence",
      목표: "goal", 어려움: "difficulty", "배운 점": "learning", "직무 연결": "coreMessage" };
    const val = (local[fieldMap[name]] || "").trim();
    return val.length >= (STEP_MIN_LENGTH[name] || 1);
  };

  const markStep = (name) => {
    const status = isStepFilled(name) ? "충분" : "보완 필요";
    const comp = { ...local.completion, [name]: status };
    setLocal(p => ({ ...p, completion: comp }));
    return comp;
  };

  const next = () => {
    const comp = markStep(step);
    if (stepIdx < steps.length - 1) { setStepIdx(stepIdx + 1); commit({ completion: comp, status: "analyzing" }); }
    else if (!showDepth) {
      commit({ completion: comp, status: "complete" });
      setShowDepth(true); setStepIdx(0);
    } else {
      commit({ completion: comp, status: "complete", depthDone: true });
      onDone(exp.id);
    }
  };
  const finishCore = () => { commit({ completion: markStep(step), status: "complete" }); onDone(exp.id); };

  const [reviewIssues, setReviewIssues] = useState(null);
  const [reviewOverall, setReviewOverall] = useState("");
  const [reviewLoading, setReviewLoading] = useState(false);
  const [reviewError, setReviewError] = useState("");

  const runConsistencyReview = async () => {
    setReviewLoading(true); setReviewError(""); setReviewIssues(null);
    try {
      const myMetrics = metrics.filter(m => m.experienceId === exp.id);
      const depthFilled = local.goal || local.difficulty || local.learning || local.jobRelevance || local.coreMessage;
      const prompt = `당신은 채용담당자 시점에서 지원자가 정리한 경험 하나를 검토합니다. 아래 단계들을 한꺼번에 보고, 서로 이어지는 이야기로서 문제가 있는지 확인하세요.

[배경] ${local.context || "(없음)"}
[문제] ${local.discoveredProblem || "(없음)"} (주어진 업무: ${local.assignedTask || "(없음)"})
[행동] ${(local.actions || []).map(a => `- (${ACTION_LABEL[a.actionType] || a.actionType}) ${a.description}`).join("\n") || "(없음)"}
[기여도] 수준: ${CONTRIB_LABEL[local.contributionLevel] || "(없음)"} / 근거: ${local.contributionEvidence || "(없음)"}
[성과] ${local.oneLineSummary || "(없음)"} / 정성 성과: ${local.qualitative || "(없음)"} / 수치: ${myMetrics.map(m => `${m.metricName} ${formatMetric(m, "exact")}`).join(", ") || "(없음)"}
${depthFilled ? `[목표] ${local.goal || "(없음)"}
[어려움] ${local.difficulty || "(없음)"}
[배운 점] ${local.learning || "(없음)"}
[직무 연결] ${local.jobRelevance || "(없음)"} / 핵심 메시지: ${local.coreMessage || "(없음)"}` : "(심화 단계는 아직 입력하지 않았습니다 — 입력된 항목만 검토하세요)"}

확인할 것:
- 빠진 정보: 특정 단계에 근거나 구체성이 없는 곳
- 개연성 문제: 앞뒤 단계가 서로 안 맞는 곳 (예: 문제에서 언급 안 된 게 성과에 갑자기 나옴, 기여도는 "혼자"라는데 행동엔 협업이 많음, 목표와 실제 행동이 안 맞음). 판단하지 말고 사실만 병치할 것.
- 구체성 부족: 숫자·장면 없이 추상적으로만 쓴 곳

없는 사실을 지어내지 마라. 각 이슈는 어느 단계(배경/문제/행동/기여도/성과/목표/어려움/배운 점/직무 연결) 얘기인지 명시하라. 입력되지 않은 심화 단계는 "빠진 정보"로 지적하지 말고 건너뛰어라 (선택 사항이므로). 문제가 없으면 issues를 빈 배열로 두라.

JSON만 응답 (마크다운 백틱 없이):
{"issues":[{"step":"성과","issue":"..."}],"overall":"전체적으로 한 줄 총평"}`;

      const res = await fetch("/api/consistency-check", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt }),
      });
      let data;
      try { data = await res.json(); }
      catch { throw new Error(`서버 응답을 읽지 못했습니다 (HTTP ${res.status})`); }
      if (!res.ok) {
        throw new Error(data?.error?.message || (typeof data?.error === "string" ? data.error : null) || `API 오류 (HTTP ${res.status}) — 응답 원문: ${JSON.stringify(data).slice(0, 300)}`);
      }
      const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("\n");
      const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
      setReviewIssues(parsed.issues || []);
      setReviewOverall(parsed.overall || "");
    } catch (e) {
      setReviewError(e.message || String(e));
    } finally {
      setReviewLoading(false);
    }
  };
  const jumpToStep = (stepName) => {
    const coreI = CORE_STEPS.indexOf(stepName);
    if (coreI >= 0) { setShowDepth(false); setStepIdx(coreI); return; }
    const depthI = DEPTH_STEPS.indexOf(stepName);
    if (depthI >= 0) { setShowDepth(true); setStepIdx(depthI); }
  };

  const fieldFor = {
    배경: [["context", "배경·상황", "언제, 어디서, 어떤 상황이었는지"], ],
    문제: [["assignedTask", "주어진 업무", "원래 맡은 업무"], ["discoveredProblem", "내가 발견한 문제", "주어진 업무와 구분해서 작성"]],
    행동: null, // 별도 UI
    기여도: null,
    성과: null,
    목표: [["goal", "목표", "수치 목표·성공 기준·제약"]],
    어려움: [["difficulty", "핵심 어려움", "부담·갈등·제약 조건"]],
    "배운 점": [["learning", "배운 점", "이전 생각 → 발견한 것 → 바뀐 행동"]],
    "직무 연결": [["jobRelevance", "직무 연결", "연결되는 직무와 보여주는 역량"], ["coreMessage", "핵심 메시지", "이 경험을 한 문장으로"]],
  };

  const aiHints = {
    배경: "팀 규모와 본인의 공식 역할이 아직 없습니다. \"몇 명 팀에서 어떤 역할이었나요?\"",
    문제: "\"어려웠다\"는 표현이 있다면 어떤 데이터·현상으로 문제라고 판단했는지 적어주세요.",
    행동: "행동을 분석·판단·실행·협업으로 나누면 면접 꼬리질문 대비가 쉬워집니다.",
    기여도: "기여 수준만 고르지 말고 근거를 함께 적으세요. \"내가 없었다면 무엇이 달라졌을까?\"",
    성과: "\"매출이 올랐다\"고 작성했다면 — 이전 기간 대비 몇 % 증가였나요? 확인할 리포트가 있나요?",
    목표: "목표를 누가 정했는지, 수치 기준이 있었는지 구분해 보세요.",
    어려움: "갈등이 있었다면 상대의 입장도 함께 기록해 두세요. 갈등 경험 문항에 재사용됩니다.",
    "배운 점": "\"협업의 중요성을 배웠다\" 같은 추상 문장 대신, 이후 실제로 바뀐 행동을 적으세요.",
    "직무 연결": "지원 직무의 JD 키워드와 연결해 보세요.",
  };

  return (
    <div style={{ maxWidth: 760 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div>
          <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>{showDepth ? "심화 분석 (선택)" : "핵심 분석"} · {exp.title}</div>
          <h2 style={{ fontSize: "var(--fs-xl)", fontWeight: 800, margin: "2px 0 0" }}>{stepIdx + 1}. {step}</h2>
        </div>
        <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
          <AutosaveIndicator state={autosave} />
          <Btn small onClick={() => { commit(); onExit(); }}>임시 저장 후 나가기</Btn>
        </div>
      </div>

      {/* 스텝 네비게이터 */}
      <div style={{ display: "flex", gap: 6, marginBottom: 16 }}>
        {steps.map((s, i) => (
          <div key={s} {...clickableProps(() => setStepIdx(i))} style={{ flex: 1, cursor: "pointer" }}>
            <div style={{ height: 4, borderRadius: "var(--r-xs)", background: i < stepIdx ? C.green : i === stepIdx ? C.blue : C.lineSoft, marginBottom: 4 }} />
            <div style={{ fontSize: "var(--fs-xs)", color: i === stepIdx ? C.text : C.faintText, fontWeight: i === stepIdx ? 700 : 500 }}>{s}</div>
          </div>
        ))}
      </div>

      <Card>
        {/* 가이드 질문 */}
        <div style={{ marginBottom: 14 }}>
          <Label>가이드 질문</Label>
          {STEP_QUESTIONS[step].map((q, i) => (
            <div key={i} style={{ fontSize: "var(--fs-base)", color: C.sub, padding: "var(--sp-1) 0" }}>· {q}</div>
          ))}
        </div>

        {/* AI 질문 힌트 */}
        <div style={{ display: "flex", gap: 8, padding: "var(--sp-4) var(--sp-5)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)", marginBottom: 16, alignItems: "flex-start" }}>
          <Badge label="AI 질문" color={C.ai} bg="#fff" />
          <span style={{ fontSize: "var(--fs-base)", color: C.text, lineHeight: 1.55 }}>{aiHints[step]}</span>
        </div>

        {/* 입력 영역 */}
        {step === "행동" && <ActionEditor local={local} setLocal={setLocal} />}
        {step === "기여도" && <ContributionEditor local={local} patch={patch} />}
        {step === "성과" && <MetricEditor expId={exp.id} metrics={metrics} setMetrics={setMetrics} local={local} patch={patch} />}
        {fieldFor[step] && fieldFor[step].map(([key, label, ph]) => (
          <div key={key} style={{ marginBottom: 12 }}>
            <Label>{label}</Label>
            <Textarea placeholder={ph} value={local[key] || ""} onChange={e => patch(key, e.target.value)} />
          </div>
        ))}

        {step === "배경" && (
          <div style={{ marginBottom: 12 }}>
            <Label>자유 메모 (참고)</Label>
            <div style={{ fontSize: "var(--fs-base)", color: C.sub, padding: "var(--sp-4) var(--sp-5)", background: C.bg, borderRadius: "var(--r-lg)", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{local.rawNote}</div>
          </div>
        )}
      </Card>

      {!isStepFilled(step) && (
        <div style={{ fontSize: "var(--fs-sm)", color: C.orangeText, marginTop: 10, background: C.orangeBg, padding: "var(--sp-3) var(--sp-5)", borderRadius: "var(--r-lg)" }}>
          이 단계 입력이 비어있거나 짧습니다. 이대로 넘어가면 "보완 필요"로 표시됩니다 — 나중에 다시 채워도 됩니다.
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14 }} className="wrap-sm">
        <Btn onClick={() => stepIdx > 0 ? setStepIdx(stepIdx - 1) : (showDepth ? (setShowDepth(false), setStepIdx(CORE_STEPS.length - 1)) : null)} disabled={stepIdx === 0 && !showDepth}>← 이전</Btn>
        <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
          {!showDepth && stepIdx === CORE_STEPS.length - 1 && (
            <>
              <Btn onClick={runConsistencyReview} disabled={reviewLoading}>{reviewLoading ? "검토 중…" : "AI로 검토받기"}</Btn>
              <Btn onClick={finishCore}>핵심 5단계로 완료</Btn>
            </>
          )}
          {showDepth && stepIdx === DEPTH_STEPS.length - 1 && (
            <Btn onClick={runConsistencyReview} disabled={reviewLoading}>{reviewLoading ? "검토 중…" : "AI로 검토받기"}</Btn>
          )}
          <Btn primary onClick={next}>
            {stepIdx < steps.length - 1 ? "다음 →" : showDepth ? "심화 분석 완료" : "심화 단계 계속 →"}
          </Btn>
        </div>
      </div>

      {((!showDepth && stepIdx === CORE_STEPS.length - 1) || (showDepth && stepIdx === DEPTH_STEPS.length - 1)) && (
        <>
          {reviewError && (
            <div style={{ marginTop: 12, padding: "var(--sp-4) var(--sp-5)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)" }}>
              <div style={{ fontSize: "var(--fs-sm)", fontWeight: 700, color: C.redText, marginBottom: 4 }}>오류</div>
              <div style={{ fontSize: "var(--fs-sm)", color: C.redText, whiteSpace: "pre-wrap", fontFamily: "monospace" }}>{reviewError}</div>
            </div>
          )}
          {reviewIssues && (
            <Card style={{ marginTop: 12, background: C.accent }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <Label>AI 검토 결과</Label>
                <span {...clickableProps(() => setReviewIssues(null), { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
              </div>
              {reviewOverall && <div style={{ fontSize: "var(--fs-base)", marginBottom: 10, lineHeight: 1.6 }}>{reviewOverall}</div>}
              {reviewIssues.length === 0 ? (
                <div style={{ fontSize: "var(--fs-base)", color: C.sub }}>뚜렷한 문제가 안 보입니다. 이대로 완료해도 좋습니다.</div>
              ) : (
                <div style={{ display: "grid", gap: 8 }}>
                  {reviewIssues.map((iss, i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.panel, border: `1px solid ${C.line}`, borderRadius: "var(--r-md)", padding: "var(--sp-4) var(--sp-5)", gap: 10 }}>
                      <div>
                        <Badge label={iss.step} color={C.blueText} bg={C.blueBg} />
                        <div style={{ fontSize: "var(--fs-base)", marginTop: 6, lineHeight: 1.5 }}>{iss.issue}</div>
                      </div>
                      <Btn small onClick={() => jumpToStep(iss.step)} style={{ flexShrink: 0 }}>이 단계로 가기</Btn>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          )}
        </>
      )}
      {!showDepth && (
        <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginTop: 10, textAlign: "right" }}>
          핵심 5단계만으로 경험 카드가 생성됩니다. 심화 4단계(목표·어려움·배운 점·직무 연결)는 나중에 추가할 수 있습니다.
        </div>
      )}
    </div>
  );
}

function ActionEditor({ local, setLocal }) {
  const actions = local.actions || [];
  const [draft, setDraft] = useState({ actionType: "analysis", description: "", isDirectAction: true, parentId: "" });
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);

  const add = () => {
    if (!draft.description.trim()) return;
    setLocal(p => ({ ...p, actions: [...(p.actions || []), {
      id: "a_" + Date.now(), actionType: draft.actionType, description: draft.description.trim(),
      isDirectAction: draft.isDirectAction, parentId: draft.parentId || null,
    }] }));
    setDraft({ actionType: "analysis", description: "", isDirectAction: true, parentId: "" });
  };

  const startEdit = (a) => { setEditingId(a.id); setEditDraft({ actionType: a.actionType, description: a.description, isDirectAction: a.isDirectAction, parentId: a.parentId || "" }); };
  const cancelEdit = () => { setEditingId(null); setEditDraft(null); };
  const saveEdit = (id) => {
    if (!editDraft.description.trim()) return;
    setLocal(p => ({ ...p, actions: p.actions.map(x => x.id === id ? {
      ...x, actionType: editDraft.actionType, description: editDraft.description.trim(),
      isDirectAction: editDraft.isDirectAction, parentId: editDraft.parentId || null,
    } : x) }));
    cancelEdit();
  };
  const remove = (id) => {
    // 부모를 지우면 자식은 삭제되지 않고 최상위로 승격됨 (데이터 손실 방지)
    setLocal(p => ({ ...p, actions: p.actions.filter(x => x.id !== id).map(x => x.parentId === id ? { ...x, parentId: null } : x) }));
  };

  const childrenOf = (id) => actions.filter(a => a.parentId === id);
  const validParentIds = new Set(actions.map(a => a.id));
  const roots = actions.filter(a => !a.parentId || !validParentIds.has(a.parentId));

  const parentOptions = (excludeId) => actions.filter(a => a.id !== excludeId);

  const renderRow = (a) => {
    const [color, bg] = ACTION_COLOR[a.actionType] || [C.blueText, C.blueBg];
    const isEditing = editingId === a.id;
    return (
      <div key={a.id}>
        {isEditing ? (
          <div style={{ padding: "var(--sp-4) 0", borderBottom: `1px solid ${C.lineSoft}` }}>
            <div style={{ display: "flex", gap: 8, marginBottom: 6, flexWrap: "wrap" }} className="wrap-sm">
              <select value={editDraft.actionType} onChange={e => setEditDraft(d => ({ ...d, actionType: e.target.value }))}
                style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-md)", border: `1px solid ${C.line}`, background: C.panel }}>
                {Object.entries(ACTION_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
              <select value={editDraft.parentId} onChange={e => setEditDraft(d => ({ ...d, parentId: e.target.value }))}
                style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-md)", border: `1px solid ${C.line}`, background: C.panel, maxWidth: 220 }}>
                <option value="">최상위 (독립 행동)</option>
                {parentOptions(a.id).map(o => <option key={o.id} value={o.id}>↳ {o.description.slice(0, 20)}{o.description.length > 20 ? "…" : ""}</option>)}
              </select>
              <label style={{ fontSize: "var(--fs-sm)", color: C.sub, display: "flex", alignItems: "center", gap: 4, whiteSpace: "nowrap" }}>
                <input type="checkbox" checked={editDraft.isDirectAction} onChange={e => setEditDraft(d => ({ ...d, isDirectAction: e.target.checked }))} /> 직접 수행
              </label>
            </div>
            <Textarea value={editDraft.description} onChange={e => setEditDraft(d => ({ ...d, description: e.target.value }))} rows={2} style={{ fontSize: "var(--fs-base)" }} />
            <div style={{ display: "flex", gap: 8, marginTop: 6 }} className="wrap-sm">
              <Btn small primary onClick={() => saveEdit(a.id)}>저장</Btn>
              <Btn small onClick={cancelEdit}>취소</Btn>
            </div>
          </div>
        ) : (
          <div style={{ display: "flex", gap: 10, alignItems: "center", padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}` }}>
            <Badge label={ACTION_LABEL[a.actionType]} color={color} bg={bg} />
            <span {...clickableProps(() => startEdit(a))} style={{ fontSize: "var(--fs-base)", flex: 1, cursor: "pointer" }}>{a.description}</span>
            {!a.isDirectAction && <Badge label="타인 수행" color={C.orangeText} bg={C.orangeBg} />}
            <span {...clickableProps(() => startEdit(a))} title="수정" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)", textDecoration: "underline" }}>수정</span>
            <span {...clickableProps(() => remove(a.id), { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
          </div>
        )}
        {childrenOf(a.id).length > 0 && (
          <div style={{ marginLeft: 10, paddingLeft: 14, borderLeft: `2px solid ${C.line}` }}>
            {childrenOf(a.id).map(renderRow)}
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={{ marginBottom: 12 }}>
      <Label>행동 카드 — 연관된 행동은 아래로 이어서 연결할 수 있습니다</Label>
      {roots.map(renderRow)}
      {actions.length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.faintText, padding: "var(--sp-3) 0" }}>아직 입력된 행동이 없습니다.</div>}

      <div style={{ marginTop: 12, padding: "var(--sp-5)", background: C.bg, borderRadius: "var(--r-lg)" }}>
        <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }} className="wrap-sm">
          <select value={draft.actionType} onChange={e => setDraft(d => ({ ...d, actionType: e.target.value }))}
            style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, background: C.panel }}>
            {Object.entries(ACTION_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <select value={draft.parentId} onChange={e => setDraft(d => ({ ...d, parentId: e.target.value }))}
            style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, background: C.panel, maxWidth: 220 }}>
            <option value="">최상위 (독립 행동)</option>
            {actions.map(a => <option key={a.id} value={a.id}>↳ {a.description.slice(0, 24)}{a.description.length > 24 ? "…" : ""}</option>)}
          </select>
          <label style={{ fontSize: "var(--fs-sm)", color: C.sub, display: "flex", alignItems: "center", gap: 4, whiteSpace: "nowrap" }}>
            <input type="checkbox" checked={draft.isDirectAction} onChange={e => setDraft(d => ({ ...d, isDirectAction: e.target.checked }))} /> 직접 수행
          </label>
        </div>
        <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
          <Input placeholder="행동 설명 — 예: 과거 3년 판매량, 장바구니 데이터 분석" value={draft.description}
            onChange={e => setDraft(d => ({ ...d, description: e.target.value }))} onKeyDown={e => e.key === "Enter" && add()} style={{ flex: 1 }} />
          <Btn small onClick={add}>추가</Btn>
        </div>
      </div>
    </div>
  );
}

function ContributionEditor({ local, patch }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <Label>기여 수준</Label>
      <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
        {Object.entries(CONTRIB_LABEL).map(([v, l]) => (
          <button key={v} onClick={() => patch("contributionLevel", v)} style={{
            fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-5)", borderRadius: "var(--r-lg)", cursor: "pointer",
            border: `1px solid ${local.contributionLevel === v ? C.text : C.line}`,
            background: local.contributionLevel === v ? C.text : C.panel, color: local.contributionLevel === v ? "#fff" : C.sub }}>
            {l}
          </button>
        ))}
      </div>
      <Label>근거 (필수 — 수준만 선택할 수 없습니다)</Label>
      <Textarea placeholder="예: 과거 3년 판매 데이터를 직접 분석하고, 사은품 제품과 구매 조건을 제안한 뒤, 프로모션 세팅과 결과 리포트까지 담당했다."
        value={local.contributionEvidence || ""} onChange={e => patch("contributionEvidence", e.target.value)} />
      <div style={{ marginTop: 10 }}>
        <Label>팀이 한 일 / 내가 한 일 구분</Label>
        <Textarea style={{ minHeight: 60 }} placeholder="팀: ... / 나: ..." value={local.personalContribution || ""} onChange={e => patch("personalContribution", e.target.value)} />
      </div>
    </div>
  );
}

function MetricEditor({ expId, metrics, setMetrics, local, patch }) {
  const mine = metrics.filter(m => m.experienceId === expId);
  const [draft, setDraft] = useState({ metricName: "", changeValue: "", unit: "", evidenceSource: "" });

  const addMetric = () => {
    if (!draft.metricName.trim() || draft.changeValue === "") return;
    setMetrics(prev => [...prev, {
      id: "m_" + Date.now() + Math.random().toString(36).slice(2, 4),
      experienceId: expId,
      metricType: "custom",
      metricName: draft.metricName.trim(),
      changeValue: Number(draft.changeValue),
      unit: draft.unit || "",
      evidenceSource: draft.evidenceSource || "",
      certainty: "needs_verification", // 직접 입력은 근거 확인 전까지 항상 "확인 필요"로 시작
      isPublic: true,
    }]);
    setDraft({ metricName: "", changeValue: "", unit: "", evidenceSource: "" });
  };
  const updateCertainty = (id, certainty) => setMetrics(prev => prev.map(m => m.id === id ? { ...m, certainty } : m));
  const removeMetric = (id) => setMetrics(prev => prev.filter(m => m.id !== id));

  return (
    <div style={{ marginBottom: 12 }}>
      <Label>성과 수치 — ExperienceMetric 단일 원본</Label>
      <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginBottom: 10, lineHeight: 1.6 }}>
        수치는 여기에만 저장됩니다. 이력서·자소서·면접 문장은 이 수치를 토큰으로 참조하며, 원본이 바뀌면 모든 문장에 반영됩니다.
        <br />"추가 확인 필요"는 자소서 AI가 이 수치를 확정적으로 쓰지 않고 조심스럽게 다루게 하고, 면접 복습 화면에서도 경고로 표시됩니다. 실제 자료로 맞는지 확인했다면 아래에서 상태를 바꿔주세요.
      </div>
      {mine.length > 0 ? mine.map(m => (
        <div key={m.id} style={{ display: "grid", gridTemplateColumns: "1fr 110px 130px 150px 20px", gap: 8, alignItems: "center", padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }} className="tbl-row">
          <span style={{ fontWeight: 600 }}>{m.metricName}</span>
          <span style={{ color: C.blueText, fontWeight: 700 }}>{formatMetric(m, "exact")}</span>
          <span style={{ fontSize: "var(--fs-sm)", color: C.sub }}>{m.comparisonBasis || m.evidenceSource || "—"}</span>
          <select value={m.certainty} onChange={e => updateCertainty(m.id, e.target.value)}
            style={{ fontFamily: font, fontSize: "var(--fs-xs)", padding: "var(--sp-2) var(--sp-2)", borderRadius: "var(--r-md)", border: `1px solid ${CERTAINTY[m.certainty][1]}55`, background: CERTAINTY[m.certainty][2], color: CERTAINTY[m.certainty][1] }}>
            {Object.entries(CERTAINTY).map(([v, [label]]) => <option key={v} value={v}>{label}</option>)}
          </select>
          <span {...clickableProps(() => removeMetric(m.id), { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
        </div>
      )) : (
        <div style={{ fontSize: "var(--fs-base)", color: C.sub, padding: "var(--sp-5)", background: C.bg, borderRadius: "var(--r-lg)" }}>
          아직 수치가 없습니다. 정량 성과가 없다면 정성 변화(CS 감소, 프로세스 표준화 등)를 아래에 적어주세요.
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 10 }} className="wrap-sm">
        <Input placeholder="지표명 (예: 매출)" value={draft.metricName} onChange={e => setDraft(d => ({ ...d, metricName: e.target.value }))} style={{ flex: 1 }} />
        <Input placeholder="변화 값 (예: 29)" value={draft.changeValue} onChange={e => setDraft(d => ({ ...d, changeValue: e.target.value }))} style={{ width: 110 }} />
        <Input placeholder="단위 (%)" value={draft.unit} onChange={e => setDraft(d => ({ ...d, unit: e.target.value }))} style={{ width: 70 }} />
        <Input placeholder="근거 자료" value={draft.evidenceSource} onChange={e => setDraft(d => ({ ...d, evidenceSource: e.target.value }))} style={{ width: 130 }} />
        <Btn small onClick={addMetric}>추가</Btn>
      </div>
      <div style={{ marginTop: 12 }}>
        <Label>정성 성과 / 성과 설명</Label>
        <Textarea style={{ minHeight: 60 }} placeholder="예: 출고 실패 0건, 협업 프로세스 표준화, 매뉴얼 배포" value={(local && local.qualitative) || ""} onChange={e => patch && patch("qualitative", e.target.value)} />
      </div>
    </div>
  );
}

/* ============================================================ 보관함 */
function Archive({ experiences, setExperiences, metrics, setMetrics, outputs, setOutputs, onOpen, onAnalyze, onGoImport, addTrash, expCategories, addExpCategory, questionBlocks, setQuestionBlocks, reviewChatHistory, setReviewChatHistory }) {
  const [showReview, setShowReview] = useState(false);
  const [view, setView] = useState("exp"); // exp | comp | question
  const [q, setQ] = useState("");
  const [filter, setFilter] = useState("all");
  const [mergeMode, setMergeMode] = useState(false);
  const [selected, setSelected] = useState([]); // 합치기 선택된 experience id들
  const [mergeStep, setMergeStep] = useState(false); // 병합 확인 화면 표시 여부
  const [collapsed, setCollapsed] = useState({}); // 카테고리 블록 접힘 상태

  const exportExcel = () => {
    const rows = experiences.map(e => ({
      제목: e.title, 카테고리: e.primaryCategory || "", 소속: e.organization || "", 역할: e.role || "",
      시작: e.startDate || "", 종료: e.endDate || "", 상태: STATUS_LABEL[e.status] || e.status,
      배경: e.context || "", 문제: e.discoveredProblem || "", 본인기여: e.personalContribution || "",
      기여근거: e.contributionEvidence || "", 성과요약: e.oneLineSummary || "",
      어려움: e.difficulty || "", 배운점: e.learning || "", 직무연결: e.jobRelevance || "",
      역량태그: (e.competencies || []).join(", "), 활용횟수: e.usageCount || 0, 최근수정: e.updatedAt || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = Object.keys(rows[0] || {}).map(k => ({ wch: Math.min(Math.max(k.length + 2, 12), 40) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "경험 목록");
    XLSX.writeFile(wb, `career-os-경험목록-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const deleteExp = (id) => {
    const exp = experiences.find(e => e.id === id);
    setExperiences(prev => prev.filter(e => e.id !== id));
    addTrash("experience", exp.title, exp);
  };

  const toggleSelect = (id) => setSelected(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
  const cancelMerge = () => { setMergeMode(false); setSelected([]); setMergeStep(false); };

  const FILTERS = [
    ["all", "전체", () => true],
    ["complete", "분석 완료", e => e.status === "complete"],
    ["needs_revision", "보완 필요", e => e.status === "needs_revision"],
    ["draft", "초기 메모", e => e.status === "draft"],
    ["unused", "아직 안 쓴 경험", e => (e.usageCount || 0) === 0],
    ["frequent", "자주 쓴 경험", e => (e.usageCount || 0) >= 3],
    ["hasMetric", "정량 성과 있음", e => metrics.some(m => m.experienceId === e.id)],
    ["needsCheck", "근거 미확인 수치", e => metrics.some(m => m.experienceId === e.id && m.certainty !== "verified")],
    ["hardship", "실패/갈등용", e => !!e.difficulty],
  ];
  const activeFilter = FILTERS.find(f => f[0] === filter) || FILTERS[0];

  const filtered = experiences.filter(e =>
    activeFilter[2](e) &&
    (q === "" || e.title.includes(q) || e.competencies.some(c => c.includes(q)) || (e.organization || "").includes(q))
  );

  const allComps = [...new Set(experiences.flatMap(e => e.competencies))]
    .sort((a, b) => experiences.filter(e => e.competencies.includes(b)).length - experiences.filter(e => e.competencies.includes(a)).length);
  const [compQ, setCompQ] = useState("");
  const [expandedComp, setExpandedComp] = useState({});
  const shownComps = allComps.filter(c => c.includes(compQ));

  const [editingBlockId, setEditingBlockId] = useState(null);
  const [noteDraft, setNoteDraft] = useState({});
  const patchBlock = (id, k, v) => setQuestionBlocks(prev => prev.map(b => b.id === id ? { ...b, [k]: v } : b));
  const addBlockExp = (id, expId) => setQuestionBlocks(prev => prev.map(b => b.id === id && expId && !b.expIds.includes(expId) ? { ...b, expIds: [...b.expIds, expId] } : b));
  const removeBlockExp = (id, expId) => setQuestionBlocks(prev => prev.map(b => b.id === id ? { ...b, expIds: b.expIds.filter(x => x !== expId) } : b));
  const addBlockNote = (id, text) => setQuestionBlocks(prev => prev.map(b => b.id === id
    ? { ...b, notes: [...(b.notes || []), { id: "n_" + Date.now() + Math.random().toString(36).slice(2, 4), text }] } : b));
  const patchBlockNote = (id, noteId, text) => setQuestionBlocks(prev => prev.map(b => b.id === id
    ? { ...b, notes: (b.notes || []).map(n => n.id === noteId ? { ...n, text } : n) } : b));
  const removeBlockNote = (id, noteId) => setQuestionBlocks(prev => prev.map(b => b.id === id
    ? { ...b, notes: (b.notes || []).filter(n => n.id !== noteId) } : b));
  const removeBlock = (id) => setQuestionBlocks(prev => prev.filter(b => b.id !== id));
  const addBlock = () => {
    const id = "qb_" + Date.now();
    setQuestionBlocks(prev => [...prev, { id, label: "", expIds: [] }]);
    setEditingBlockId(id);
  };

  return (
    <div>
      {showReview ? (
        <EssayChat
          title="AI 경험 진단"
          subtitle="현직 채용담당자 시점으로 전체 경험을 검토합니다"
          systemPrompt={EXPERIENCE_REVIEW_SYSTEM_PROMPT}
          contextText={buildReviewContext(experiences, metrics)}
          autoStartMessage="제 경험 데이터를 전체적으로 검토하고, 우선순위 높은 보완점부터 짚어주세요."
          inputPlaceholder="특정 경험에 대해 더 물어보거나, 다른 관점으로 다시 봐달라고 요청해보세요"
          onClose={() => setShowReview(false)}
          history={reviewChatHistory}
          onHistoryChange={setReviewChatHistory}
        />
      ) : mergeStep ? (
        <MergeReview ids={selected} experiences={experiences} setExperiences={setExperiences}
          metrics={metrics} setMetrics={setMetrics} outputs={outputs} setOutputs={setOutputs}
          addTrash={addTrash} onDone={(id) => { cancelMerge(); onOpen(id); }} onCancel={() => setMergeStep(false)} />
      ) : (
      <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }} className="wrap-sm">
        <H2>경험 보관함</H2>
        <div style={{ display: "flex", gap: 6 }} className="wrap-sm">
          {mergeMode ? (
            <Btn small onClick={cancelMerge}>선택 모드 종료</Btn>
          ) : (
            <>
              <Btn small onClick={onGoImport}>파일 가져오기</Btn>
              <Btn small onClick={exportExcel}>엑셀로 내보내기</Btn>
              <Btn small onClick={() => setShowReview(true)}>AI 진단 받기</Btn>
              <Btn small onClick={() => setMergeMode(true)}>선택 모드</Btn>
              {[["exp", "경험별"], ["comp", "역량별"], ["question", "질문별"]].map(([v, l]) => (
                <Btn key={v} small primary={view === v} onClick={() => setView(v)}>{l}</Btn>
              ))}
            </>
          )}
        </div>
      </div>

      {mergeMode && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)", padding: "var(--sp-4) var(--sp-5)", marginBottom: 14, flexWrap: "wrap", gap: 8 }} className="wrap-sm">
          <span style={{ fontSize: "var(--fs-base)" }}>{selected.length}개 선택됨</span>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }} className="wrap-sm">
            <CategorySelect value="" options={expCategories} placeholder="카테고리 일괄 지정"
              onAddOption={addExpCategory}
              onChange={(v) => { setExperiences(prev => prev.map(e => selected.includes(e.id) ? { ...e, primaryCategory: v } : e)); }} />
            <Btn small disabled={selected.length === 0} onClick={() => {
              if (!window.confirm(`선택한 경험 ${selected.length}개를 삭제할까요?`)) return;
              selected.forEach(id => { const e = experiences.find(x => x.id === id); if (e) addTrash("experience", e.title, e); });
              setExperiences(prev => prev.filter(e => !selected.includes(e.id)));
              setSelected([]);
            }}>선택 삭제</Btn>
            <Btn small primary disabled={selected.length < 2} onClick={() => setMergeStep(true)}>선택한 경험 합치기 →</Btn>
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }} className="wrap-sm">
        <Input placeholder="경험·역량·소속 검색" value={q} onChange={e => setQ(e.target.value)} style={{ maxWidth: 300 }} />
        {FILTERS.map(([v, l]) => (
          <button key={v} onClick={() => setFilter(v)} style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-5)", borderRadius: "var(--r-lg)", cursor: "pointer", whiteSpace: "nowrap",
            border: `1px solid ${filter === v ? C.text : C.line}`, background: filter === v ? C.text : C.panel, color: filter === v ? "#fff" : C.sub }}>{l}</button>
        ))}
      </div>

      {view === "exp" && (() => {
        const groups = {};
        filtered.forEach(e => {
          const key = e.primaryCategory || "미분류";
          (groups[key] = groups[key] || []).push(e);
        });
        const orderedKeys = [...expCategories.filter(c => groups[c]), ...Object.keys(groups).filter(k => k !== "미분류" && !expCategories.includes(k)), ...(groups["미분류"] ? ["미분류"] : [])];

        const renderCard = (e) => {
          const doneCnt = Object.values(e.completion).filter(v => v === "충분").length;
          return (
            <Card key={e.id} onClick={() => mergeMode ? toggleSelect(e.id) : onOpen(e.id)}
              style={mergeMode && selected.includes(e.id) ? { borderColor: C.text, background: C.bg } : {}}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  {mergeMode && <input type="checkbox" checked={selected.includes(e.id)} onChange={() => toggleSelect(e.id)} onClick={ev => ev.stopPropagation()} />}
                  <div style={{ fontWeight: 700, fontSize: "var(--fs-md)" }}>{e.title}</div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <Badge label={STATUS_LABEL[e.status]} color={STATUS_COLOR[e.status][0]} bg={STATUS_COLOR[e.status][1]} />
                  {!mergeMode && <span {...clickableProps(ev => { ev.stopPropagation(); deleteExp(e.id); }, { label: "닫기" })}
                    title="삭제 (휴지통에서 복구 가능)" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>}
                </div>
              </div>
              <div style={{ fontSize: "var(--fs-sm)", color: C.sub, marginBottom: 8 }}>{e.organization} · {e.startDate}~{e.endDate}</div>
              {e.oneLineSummary && <div style={{ fontSize: "var(--fs-base)", lineHeight: 1.55, marginBottom: 10, color: C.text }}>{e.oneLineSummary}</div>}
              <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 10 }}>
                {e.competencies.slice(0, 3).map(c => <Badge key={c} label={"#" + c} color={C.sub} bg={C.lineSoft} />)}
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: "var(--fs-xs)", color: C.faintText, marginBottom: 8 }}>
                <span>완성도 {doneCnt}/9 {e.status === "complete" && !e.depthDone && "· 심화 미입력"}</span>
                <span>활용 {e.usageCount}회 · {e.updatedAt}</span>
              </div>
              <div {...clickableProps(ev => ev.stopPropagation())} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>카테고리:</span>
                <CategorySelect value={e.primaryCategory} options={expCategories} placeholder="미분류"
                  onAddOption={addExpCategory}
                  onChange={(v) => setExperiences(prev => prev.map(x => x.id === e.id ? { ...x, primaryCategory: v } : x))} />
              </div>
              {e.status !== "complete" && (
                <div style={{ marginTop: 10 }}><Btn small onClick={ev => { ev.stopPropagation(); onAnalyze(e.id); }}>분석 이어하기 →</Btn></div>
              )}
            </Card>
          );
        };

        return orderedKeys.map(key => (
          <div key={key} style={{ marginBottom: 20 }}>
            <div {...clickableProps(() => setCollapsed(p => ({ ...p, [key]: !p[key] })))} style={{
              display: "flex", alignItems: "center", gap: 8, cursor: "pointer", marginBottom: 10 }}>
              <span style={{ fontSize: "var(--fs-2xs)", color: C.faintText, transform: collapsed[key] ? "rotate(-90deg)" : "rotate(0deg)" }}>▾</span>
              <span style={{ fontSize: "var(--fs-base)", fontWeight: 700 }}>{key}</span>
              <span style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>({groups[key].length})</span>
            </div>
            {!collapsed[key] && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }} className="stack-sm">
                {groups[key].map(renderCard)}
              </div>
            )}
          </div>
        ));
      })()}

      {view === "comp" && (
        <div>
          <Input placeholder="역량 검색" value={compQ} onChange={e => setCompQ(e.target.value)} style={{ maxWidth: 260, marginBottom: 14 }} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }} className="stack-sm">
            {shownComps.map(c => {
              const hits = experiences.filter(e => e.competencies.includes(c));
              const isOpen = expandedComp[c];
              const shown = isOpen ? hits : hits.slice(0, 4);
              return (
                <Card key={c}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <div style={{ fontWeight: 700, fontSize: "var(--fs-base)" }}>#{c}</div>
                    <span style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>{hits.length}개</span>
                  </div>
                  {shown.map(e => (
                    <div key={e.id} {...clickableProps(() => onOpen(e.id))} style={{ padding: "var(--sp-2) 0", borderBottom: `1px solid ${C.lineSoft}`, cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                        <span style={{ fontSize: "var(--fs-base)", fontWeight: 600 }}>{e.title}</span>
                        <Badge label={STATUS_LABEL[e.status]} color={STATUS_COLOR[e.status][0]} bg={STATUS_COLOR[e.status][1]} />
                      </div>
                      {e.oneLineSummary && <div style={{ fontSize: "var(--fs-xs)", color: C.faintText, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.oneLineSummary}</div>}
                    </div>
                  ))}
                  {hits.length > 4 && (
                    <div {...clickableProps(() => setExpandedComp(p => ({ ...p, [c]: !p[c] })))} style={{ fontSize: "var(--fs-sm)", color: C.sub, cursor: "pointer", marginTop: 6 }}>
                      {isOpen ? "접기" : `+ ${hits.length - 4}개 더보기`}
                    </div>
                  )}
                </Card>
              );
            })}
          </div>
          {shownComps.length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.faintText }}>일치하는 역량이 없습니다.</div>}
        </div>
      )}

      {view === "question" && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }} className="stack-sm">
            {questionBlocks.map(b => {
              const hits = b.expIds.map(id => experiences.find(e => e.id === id)).filter(Boolean);
              const notes = b.notes || [];
              const isEditing = editingBlockId === b.id;
              const candidates = experiences.filter(e => !b.expIds.includes(e.id));
              const total = hits.length + notes.length;
              return (
                <Card key={b.id}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, gap: 8 }} className="wrap-sm">
                    {isEditing ? (
                      <Input value={b.label} placeholder="질문 내용" onChange={e => patchBlock(b.id, "label", e.target.value)}
                        style={{ fontWeight: 700, fontSize: "var(--fs-base)", border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
                    ) : (
                      <div {...clickableProps(() => setEditingBlockId(b.id))} style={{ fontWeight: 700, fontSize: "var(--fs-base)", cursor: "pointer", flex: 1 }}>{b.label || "(제목 없음 — 클릭해서 입력)"}</div>
                    )}
                    <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                      {total === 0 ? <Badge label="준비 부족" color={C.redText} bg={C.redBg} /> : <span style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>{total}개</span>}
                      {isEditing ? (
                        <Btn small onClick={() => setEditingBlockId(null)}>완료</Btn>
                      ) : (
                        <span {...clickableProps(() => removeBlock(b.id), { label: "닫기" })} title="질문 삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                      )}
                    </div>
                  </div>

                  {hits.map(e => (
                    <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "var(--sp-2) 0", borderBottom: `1px solid ${C.lineSoft}` }}>
                      <Badge label="경험" color={C.sub} bg={C.lineSoft} />
                      <span {...clickableProps(() => !isEditing && onOpen(e.id))} style={{ fontSize: "var(--fs-base)", fontWeight: 600, flex: 1, cursor: isEditing ? "default" : "pointer" }}>{e.title}</span>
                      {isEditing && <span {...clickableProps(() => removeBlockExp(b.id, e.id), { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>}
                    </div>
                  ))}

                  {notes.map(n => (
                    <div key={n.id} style={{ display: "flex", alignItems: "flex-start", gap: 8, padding: "var(--sp-2) 0", borderBottom: `1px solid ${C.lineSoft}` }}>
                      <Badge label="메모" color={C.sub} bg={C.lineSoft} />
                      {isEditing ? (
                        <Textarea value={n.text} onChange={e => patchBlockNote(b.id, n.id, e.target.value)} style={{ flex: 1, minHeight: 44, fontSize: "var(--fs-base)" }} />
                      ) : (
                        <span style={{ fontSize: "var(--fs-base)", lineHeight: 1.5, flex: 1, whiteSpace: "pre-wrap" }}>{n.text}</span>
                      )}
                      {isEditing && <span {...clickableProps(() => removeBlockNote(b.id, n.id), { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>}
                    </div>
                  ))}

                  {total === 0 && !isEditing && <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>이 질문에 쓸 경험이나 메모가 아직 없습니다.</div>}

                  {isEditing && (
                    <div style={{ marginTop: 10, display: "grid", gap: 8 }}>
                      <select value="" onChange={e => addBlockExp(b.id, e.target.value)}
                        style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-md)", border: `1px solid ${C.line}`, background: C.panel, width: "100%" }}>
                        <option value="">+ 경험 불러오기</option>
                        {candidates.map(e => <option key={e.id} value={e.id}>{e.title}</option>)}
                      </select>
                      <div style={{ display: "flex", gap: 6 }} className="wrap-sm">
                        <Textarea placeholder="경험 없이 바로 메모나 답변 초안을 적어도 됩니다" value={noteDraft[b.id] || ""}
                          onChange={e => setNoteDraft(p => ({ ...p, [b.id]: e.target.value }))} style={{ flex: 1, minHeight: 44, fontSize: "var(--fs-base)" }} />
                        <Btn small onClick={() => { if ((noteDraft[b.id] || "").trim()) { addBlockNote(b.id, noteDraft[b.id].trim()); setNoteDraft(p => ({ ...p, [b.id]: "" })); } }}>메모 추가</Btn>
                      </div>
                    </div>
                  )}
                </Card>
              );
            })}
          </div>
          <div style={{ marginTop: 12 }}><Btn small onClick={addBlock}>+ 질문 추가</Btn></div>
        </div>
      )}
      </>
      )}
    </div>
  );
}

/* ============================================================ 경험 합치기 검토 */
function MergeReview({ ids, experiences, setExperiences, metrics, setMetrics, outputs, setOutputs, addTrash, onDone, onCancel }) {
  const items = experiences.filter(e => ids.includes(e.id));
  const [primaryId, setPrimaryId] = useState(ids[0]);
  const [mergedNote, setMergedNote] = useState(
    items.map(e => `[${e.title}]\n${e.rawNote || e.oneLineSummary || "(내용 없음)"}`).join("\n\n")
  );

  const confirm = () => {
    const primary = items.find(e => e.id === primaryId);
    const others = items.filter(e => e.id !== primaryId);
    const otherIds = others.map(e => e.id);

    const merged = {
      ...primary,
      rawNote: mergedNote,
      competencies: [...new Set(items.flatMap(e => e.competencies))],
      tags: [...new Set(items.flatMap(e => e.tags))],
      actions: items.flatMap(e => e.actions),
    };

    setMetrics(prev => prev.map(m => otherIds.includes(m.experienceId) ? { ...m, experienceId: primaryId } : m));
    setOutputs(prev => prev.map(o => otherIds.includes(o.experienceId) ? { ...o, experienceId: primaryId } : o));
    setExperiences(prev => prev.filter(e => !otherIds.includes(e.id)).map(e => e.id === primaryId ? merged : e));
    others.forEach(o => addTrash("experience", o.title, o));

    onDone(primaryId);
  };

  return (
    <div style={{ maxWidth: 700 }}>
      <H2>경험 합치기</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.6 }}>
        기준이 될 경험을 하나 고르세요. 나머지는 휴지통으로 이동하고, 그 경험에 딸린 성과 수치·활용 문장은 기준 경험으로 옮겨집니다.
      </div>
      {items.map(e => (
        <label key={e.id} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "var(--sp-4) 0", borderBottom: `1px solid ${C.lineSoft}`, cursor: "pointer" }}>
          <input type="radio" name="primary" checked={primaryId === e.id} onChange={() => setPrimaryId(e.id)} style={{ marginTop: 3 }} />
          <div>
            <div style={{ fontWeight: 700, fontSize: "var(--fs-base)", display: "flex", gap: 6, alignItems: "center" }}>
              {e.title} {primaryId === e.id && <Badge label="기준" color={C.greenText} bg={C.greenBg} />}
            </div>
            <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>{e.organization} · {e.startDate}~{e.endDate}</div>
          </div>
        </label>
      ))}
      <div style={{ marginTop: 16 }}>
        <Label>병합된 메모 (수정 가능)</Label>
        <Textarea rows={9} value={mergedNote} onChange={e => setMergedNote(e.target.value)} />
      </div>
      <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginTop: 8, lineHeight: 1.6 }}>
        배경·문제·행동 등 세부 필드는 기준 경험의 내용이 유지됩니다. 합친 뒤 경험 분석에서 전체 내용을 다시 확인·정리하는 것을 권장합니다.
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 16 }} className="wrap-sm">
        <Btn primary onClick={confirm}>합치기 확정</Btn>
        <Btn onClick={onCancel}>취소</Btn>
      </div>
    </div>
  );
}

/* ============================================================ 경험 상세 */
function ExperienceDetail({ exp, metrics, setMetrics, outputs, setOutputs, setExperiences, onBack, onAnalyze, onDeleted, addTrash }) {
  const [tab, setTab] = useState("요약");
  const mine = metrics.filter(m => m.experienceId === exp.id);
  const myOutputs = outputs.filter(o => o.experienceId === exp.id);

  const patchField = (k, v) => setExperiences(prev => prev.map(e => e.id === exp.id ? { ...e, [k]: v } : e));
  const setExpLocal = (updater) => setExperiences(prev => prev.map(e => e.id === exp.id ? (typeof updater === "function" ? updater(e) : updater) : e));

  const approve = (id) => setOutputs(prev => prev.map(o => o.id === id ? { ...o, approvalStatus: "approved", isStale: false } : o));
  const reject = (id) => setOutputs(prev => prev.map(o => o.id === id ? { ...o, approvalStatus: "rejected" } : o));
  const deleteExp = () => {
    setExperiences(prev => prev.filter(e => e.id !== exp.id));
    addTrash("experience", exp.title, exp);
    onDeleted();
  };

  const [tagDraft, setTagDraft] = useState("");
  const addTag = () => {
    const t = tagDraft.trim();
    if (!t || exp.competencies.includes(t)) return;
    setExperiences(prev => prev.map(e => e.id === exp.id ? { ...e, competencies: [...e.competencies, t] } : e));
    setTagDraft("");
  };
  const removeTag = (t) => setExperiences(prev => prev.map(e => e.id === exp.id ? { ...e, competencies: e.competencies.filter(c => c !== t) } : e));

  const [chatMode, setChatMode] = useState(null); // { type: "regenerate", outputId } | { type: "new" }
  const [showReview, setShowReview] = useState(false);
  const [reviewHistory, setReviewHistory] = useState([]);
  const expApp = { company: exp.organization || exp.title, position: exp.role || "", requirements: [] };
  const closeChat = () => setChatMode(null);

  const missing = [];
  if (!exp.context) missing.push("배경·상황 정보가 없습니다.");
  if (!exp.contributionEvidence) missing.push("성과가 본인의 행동 때문이라는 근거가 필요합니다.");
  if (!exp.difficulty) missing.push("실패·아쉬움 정보가 없어 실패 경험 문항에 활용할 수 없습니다.");

  return (
    <div style={{ maxWidth: 820 }}>
      <div {...clickableProps(onBack)} style={{ fontSize: "var(--fs-base)", color: C.sub, cursor: "pointer", marginBottom: 10 }}>← 경험 보관함</div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }} className="wrap-sm">
        <h2 style={{ fontSize: "var(--fs-2xl)", fontWeight: 800, margin: 0 }}>{exp.title}</h2>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }} className="wrap-sm">
          <Badge label={STATUS_LABEL[exp.status]} color={STATUS_COLOR[exp.status][0]} bg={STATUS_COLOR[exp.status][1]} />
          <Btn small onClick={() => setShowReview(true)}>이 경험만 AI 진단</Btn>
          <Btn small onClick={() => onAnalyze(exp.id)}>{exp.depthDone ? "수정하기" : "심화 분석 계속"}</Btn>
          <span {...clickableProps(deleteExp, { label: "닫기" })} title="이 경험 삭제 (휴지통에서 복구 가능)" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-md)", padding: "0 var(--sp-2)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
        </div>
      </div>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16 }}>{exp.organization} · {exp.role} · {exp.startDate}~{exp.endDate}</div>

      {showReview && (
        <div style={{ marginBottom: 18 }}>
          <EssayChat
            title="AI 경험 진단 (이 경험만)"
            subtitle={exp.title}
            systemPrompt={EXPERIENCE_REVIEW_SYSTEM_PROMPT}
            contextText={buildReviewContext([exp], metrics)}
            autoStartMessage="이 경험 하나만 자세히 검토하고, 우선순위 높은 보완점부터 짚어주세요."
            inputPlaceholder="더 물어보거나, 다른 관점으로 다시 봐달라고 요청해보세요"
            onClose={() => setShowReview(false)}
            history={reviewHistory}
            onHistoryChange={setReviewHistory}
          />
        </div>
      )}

      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.line}`, marginBottom: 18 }}>
        {["요약", "사실", "행동", "성과", "활용 문장", "완성도"].map(t => (
          <div key={t} {...clickableProps(() => setTab(t))} style={{ padding: "var(--sp-3) var(--sp-5)", fontSize: "var(--fs-base)", fontWeight: tab === t ? 700 : 500, cursor: "pointer",
            color: tab === t ? C.text : C.sub, borderBottom: tab === t ? `2px solid ${C.text}` : "2px solid transparent", marginBottom: -1 }}>{t}</div>
        ))}
      </div>

      {tab === "요약" && (
        <div style={{ display: "grid", gap: 12 }}>
          <Card><Label>한 줄 요약</Label><Textarea value={exp.oneLineSummary || ""} placeholder="이 경험을 한 줄로 요약하면" onChange={e => patchField("oneLineSummary", e.target.value)} style={{ fontSize: "var(--fs-md)", minHeight: 50 }} /></Card>
          <Card style={{ background: C.accent }}><Label>핵심 메시지</Label><Textarea value={exp.coreMessage || ""} placeholder="이 경험의 핵심 메시지" onChange={e => patchField("coreMessage", e.target.value)} style={{ fontSize: "var(--fs-md)", minHeight: 50, background: "transparent" }} /></Card>
          <Card>
            <Label>대표 역량 — 추가·삭제 가능</Label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
              {exp.competencies.map(c => (
                <span key={c} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                  <Badge label={"#" + c} color={C.blueText} bg={C.blueBg} />
                  <span {...clickableProps(() => removeTag(c), { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-xs)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                </span>
              ))}
              {exp.competencies.length === 0 && <span style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>아직 태그가 없습니다.</span>}
            </div>
            <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
              <Input placeholder="역량 태그 추가 (예: 협상력)" value={tagDraft} onChange={e => setTagDraft(e.target.value)}
                onKeyDown={e => e.key === "Enter" && addTag()} style={{ maxWidth: 220 }} />
              <Btn small onClick={addTag}>추가</Btn>
            </div>
          </Card>
          {missing.length > 0 && (
            <Card style={{ background: C.accent }}>
              <Label>부족한 정보</Label>
              {missing.map((m, i) => <div key={i} style={{ fontSize: "var(--fs-base)", color: C.orangeText, padding: "var(--sp-1) 0" }}>· {m}</div>)}
            </Card>
          )}
        </div>
      )}

      {tab === "사실" && (
        <Card>
          <Label>사실 보관함 — 직접 수정 가능</Label>
          {[["organization", "소속"], ["role", "역할"], ["assignedTask", "주어진 업무"], ["discoveredProblem", "발견한 문제"]].map(([k, label]) => (
            <div key={k} style={{ display: "grid", gridTemplateColumns: "110px 1fr", gap: 10, padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)", alignItems: "center" }}>
              <span style={{ color: C.faintText, fontWeight: 600 }}>{label}</span>
              <Input value={exp[k] || ""} placeholder="미입력" onChange={e => patchField(k, e.target.value)} style={{ border: "none", padding: "var(--sp-1) 0" }} />
            </div>
          ))}
          <div style={{ display: "grid", gridTemplateColumns: "110px 1fr", gap: 10, padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)", alignItems: "center" }}>
            <span style={{ color: C.faintText, fontWeight: 600 }}>기여 수준</span>
            <select value={exp.contributionLevel || ""} onChange={e => patchField("contributionLevel", e.target.value)}
              style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, background: C.panel, width: 200 }}>
              <option value="">미입력</option>
              {Object.entries(CONTRIB_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
          </div>
          <div style={{ padding: "var(--sp-3) 0", fontSize: "var(--fs-base)" }}>
            <div style={{ color: C.faintText, fontWeight: 600, marginBottom: 6 }}>기여 근거</div>
            <Textarea value={exp.contributionEvidence || ""} placeholder="본인 기여를 증명할 근거를 적어주세요" onChange={e => patchField("contributionEvidence", e.target.value)} />
          </div>
          <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginTop: 10 }}>여기서 고친 내용은 활용 문장이 참조하는 원본에 바로 반영됩니다. 이미 승인된 문장 자체는 자동으로 바뀌지 않으니, 필요하면 "재생성"으로 새로 만들어주세요.</div>
        </Card>
      )}

      {tab === "행동" && (
        <Card>
          <ActionEditor local={exp} setLocal={setExpLocal} />
        </Card>
      )}

      {tab === "성과" && (
        <Card>
          <MetricEditor expId={exp.id} metrics={metrics} setMetrics={setMetrics} local={exp} patch={patchField} />
        </Card>
      )}

      {tab === "활용 문장" && chatMode && (
        <EssayChat
          title={chatMode.type === "regenerate" ? "문장 재생성" : "새 이력서 문장 만들기"}
          subtitle={exp.title}
          systemPrompt={ESSAY_COACH_SYSTEM_PROMPT}
          contextText={buildEssayContext(expApp,
            chatMode.type === "regenerate"
              ? { question: `아래 기존 문장을 더 설득력 있게 다시 써주세요 (성과 중심, 이력서용 한 문장):\n"${outputs.find(o => o.id === chatMode.outputId)?.content || ""}"`, characterLimit: 150 }
              : { question: "이 경험을 바탕으로 이력서에 쓸 성과 중심의 한 문장을 만들어주세요.", characterLimit: 150 },
            [exp], metrics)}
          autoStartMessage={chatMode.type === "regenerate" ? "다시 써주세요." : "문장을 만들어주세요."}
          onClose={closeChat}
          onSaveDraft={(text) => {
            if (chatMode.type === "regenerate") {
              setOutputs(prev => prev.map(o => o.id === chatMode.outputId
                ? { ...o, content: text, version: o.version + 1, approvalStatus: "ai_draft", isStale: false } : o));
            } else {
              setOutputs(prev => [...prev, {
                id: "o_" + Date.now(), experienceId: exp.id, outputType: "resume",
                content: text, referencedMetricIds: [], version: 1, isAiGenerated: true,
                approvalStatus: "ai_draft", isStale: false,
              }]);
            }
            closeChat();
          }}
          saveDraftLabel="이 문장으로 저장"
        />
      )}

      {tab === "활용 문장" && !chatMode && (
        <div style={{ display: "grid", gap: 12 }}>
          <div style={{ fontSize: "var(--fs-sm)", color: C.sub, lineHeight: 1.6 }}>
            수치는 <span style={{ background: C.greenBg, color: C.greenText, padding: "var(--sp-1) var(--sp-2)", borderRadius: "var(--r-md)", fontWeight: 600 }}>토큰</span>으로 원본을 참조합니다. 승인된 문장만 제출본에 사용할 수 있습니다.
          </div>
          {myOutputs.filter(o => o.approvalStatus !== "rejected").map(o => {
            const ap = APPROVAL[o.approvalStatus];
            return (
              <Card key={o.id} style={o.isStale ? { borderColor: C.red } : {}}>
                <div style={{ display: "flex", gap: 6, marginBottom: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <Badge label={{ resume: "이력서", essay: "자소서", interview: "면접", portfolio: "포트폴리오" }[o.outputType]} color={C.sub} bg={C.lineSoft} />
                  {o.style && <Badge label={{ role_focused: "역할 중심", result_focused: "성과 중심", competency_focused: "역량 중심" }[o.style]} color={C.sub} bg={C.lineSoft} />}
                  {o.isAiGenerated && <Badge label="AI 생성" color={C.ai} bg={C.aiBg} />}
                  <Badge label={ap.label} color={ap.color} bg={ap.bg} />
                  {o.isStale && <Badge label="참조 수치 변경됨 · 재승인 필요" color={C.redText} bg={C.redBg} />}
                  <span style={{ fontSize: "var(--fs-xs)", color: C.faintText, marginLeft: "auto" }}>v{o.version}</span>
                </div>
                <div style={{ fontSize: "var(--fs-md)", lineHeight: 1.65 }}><TokenText text={o.content} metrics={metrics} /></div>
                <div style={{ display: "flex", gap: 8, marginTop: 12 }} className="wrap-sm">
                  {o.approvalStatus !== "approved" && <Btn small primary onClick={() => approve(o.id)}>승인</Btn>}
                  {o.isStale && <Btn small primary onClick={() => approve(o.id)}>확인 후 재승인</Btn>}
                  <Btn small onClick={() => setChatMode({ type: "regenerate", outputId: o.id })}>재생성</Btn>
                  <Btn small onClick={() => reject(o.id)}>폐기</Btn>
                </div>
              </Card>
            );
          })}
          <div><Btn onClick={() => setChatMode({ type: "new" })}>+ 이력서 문장 만들기</Btn></div>
        </div>
      )}

      {tab === "완성도" && (
        <Card>
          <Label>항목별 완성도 — 점수보다 보완할 항목이 우선</Label>
          {Object.entries(exp.completion).map(([k, v]) => (
            <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }}>
              <span>{k}{DEPTH_STEPS.includes(k) && <span style={{ fontSize: "var(--fs-xs)", color: C.faintText, marginLeft: 6 }}>심화</span>}</span>
              <Badge label={v} color={v === "충분" ? C.greenText : v === "보완 필요" ? C.orangeText : C.faintText} bg={v === "충분" ? C.greenBg : v === "보완 필요" ? C.orangeBg : C.lineSoft} />
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}

/* ============================================================ 파일 가져오기 (AI 추출 → 검토 → 반영) */
function ImportFlow({ setExperiences, setSkills, setCerts, setResumeProfile, onDone, experiences }) {
  const [phase, setPhase] = useState("input"); // input | loading | review | done
  const [raw, setRaw] = useState("");
  const [fileName, setFileName] = useState("");
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [workbook, setWorkbook] = useState(null); // 여러 시트가 있을 때 선택 UI용

  const loadSheet = (wb, sheetName) => {
    const ws = wb.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(ws);
    setRaw(csv);
    setWorkbook(null);
  };

  const onFile = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFileName(f.name); setError(""); setWorkbook(null);
    if (/\.(txt|md)$/i.test(f.name)) {
      const r = new FileReader();
      r.onload = () => setRaw(String(r.result));
      r.readAsText(f);
    } else if (/\.docx$/i.test(f.name)) {
      const r = new FileReader();
      r.onload = async () => {
        try {
          const { value } = await mammoth.extractRawText({ arrayBuffer: r.result });
          setRaw(value);
        } catch {
          setError("Word 파일을 읽지 못했습니다. 내용을 복사해 아래에 붙여넣어 주세요.");
        }
      };
      r.readAsArrayBuffer(f);
    } else if (/\.(xlsx|xls|csv)$/i.test(f.name)) {
      const r = new FileReader();
      r.onload = () => {
        try {
          const wb = XLSX.read(r.result, { type: "array" });
          if (wb.SheetNames.length === 1) {
            loadSheet(wb, wb.SheetNames[0]);
          } else {
            setWorkbook(wb); // 여러 시트 → 선택 UI 표시
          }
        } catch {
          setError("엑셀 파일을 읽지 못했습니다. 파일이 손상되지 않았는지 확인해 주세요.");
        }
      };
      r.readAsArrayBuffer(f);
    } else {
      setError("PDF는 이 프로토타입에서 직접 읽지 못합니다. 내용을 복사해 아래에 붙여넣어 주세요. (2차 구현에서 지원)");
    }
  };

  const extract = async () => {
    if (!raw.trim()) return;
    setPhase("loading"); setError("");
    try {
      const res = await fetch("/api/extract", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ raw }),
      });

      let data;
      try { data = await res.json(); }
      catch { throw new Error(`서버 응답을 읽지 못했습니다 (HTTP ${res.status})`); }

      if (!res.ok) {
        const msg = data?.error?.message
          || (typeof data?.error === "string" ? data.error : null)
          || `API 오류 (HTTP ${res.status}) — 응답 원문: ${JSON.stringify(data).slice(0, 500)}`;
        throw new Error(msg);
      }

      const textBlocks = (data.content || []).filter(b => b.type === "text").map(b => b.text);
      if (textBlocks.length === 0) {
        throw new Error("API 응답에 텍스트 내용이 없습니다. 응답: " + JSON.stringify(data).slice(0, 300));
      }
      const text = textBlocks.join("\n");

      let parsed;
      try {
        parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
      } catch (e) {
        throw new Error("AI 응답을 JSON으로 해석하지 못했습니다: " + e.message + "\n\n응답 원문 일부: " + text.slice(0, 300));
      }

      setResult(normalize(parsed));
      setPhase("review");
    } catch (e) {
      setError(e.message || String(e));
      setPhase("input");
    }
  };

  const isDuplicate = (title, organization) => experiences.some(e =>
    e.title.trim().toLowerCase() === (title || "").trim().toLowerCase() &&
    (e.organization || "").trim().toLowerCase() === (organization || "").trim().toLowerCase()
  );

  const normalize = (r) => ({
    experiences: (r.experiences || []).map((e, i) => {
      const dup = isDuplicate(e.title, e.organization);
      return { ...e, _id: "ix" + i, _include: !dup, _duplicate: dup };
    }),
    skills: (r.skills || []).map((s, i) => ({ ...s, _id: "is" + i, _include: true })),
    certs: (r.certs || []).map((c, i) => ({ ...c, _id: "ic" + i, _include: true })),
    profile: r.profile || {},
    _profileInclude: !!(r.profile && (r.profile.name || r.profile.email)),
  });

  const patchProfile = (k, v) => setResult(p => ({ ...p, profile: { ...p.profile, [k]: v } }));

  const patchExp = (id, k, v) => setResult(p => ({ ...p, experiences: p.experiences.map(e => e._id === id ? { ...e, [k]: v } : e) }));
  const patchMetric = (eid, mi, k, v) => setResult(p => ({ ...p, experiences: p.experiences.map(e => e._id === eid ? { ...e, metrics: e.metrics.map((m, i) => i === mi ? { ...m, [k]: v } : m) } : e) }));
  const toggle = (list, id) => setResult(p => ({ ...p, [list]: p[list].map(x => x._id === id ? { ...x, _include: !x._include } : x) }));

  const commit = () => {
    const now = "2026-07-21";
    const newExps = result.experiences.filter(e => e._include).map(e => ({
      id: "e_" + Date.now() + Math.random().toString(36).slice(2, 6),
      title: e.title, organization: e.organization, role: e.role, experienceType: e.experienceType || "other",
      startDate: e.startDate, endDate: e.endDate, rawNote: e.rawNote,
      context: "", assignedTask: "", discoveredProblem: "", goal: "", personalContribution: "",
      contributionLevel: "", contributionEvidence: "", coreMessage: "", oneLineSummary: "",
      status: "draft", depthDone: false, usageCount: 0, updatedAt: now, primaryCategory: "",
      competencies: e.competencies || [], tags: ["가져옴"], actions: [],
      importedMetrics: e.metrics, // 분석 성과 단계에서 확인 후 정식 Metric으로 승격
      completion: Object.fromEntries([...CORE_STEPS, ...DEPTH_STEPS].map(s => [s, "미입력"])),
    }));
    setExperiences(p => [...p, ...newExps]);
    setSkills(p => {
      const names = p.map(x => x.name.toLowerCase());
      const add = result.skills.filter(s => s._include && !names.includes(s.name.toLowerCase()))
        .map(s => ({ id: "s_" + Date.now() + Math.random().toString(36).slice(2, 4), name: s.name, category: s.category || "tool", summary: "가져온 항목 — 활용 범위 확인 필요", scopeItems: (s.scopeItems || []).map((it, i) => ({ id: "sc_i" + i + Date.now(), ...it })) }));
      return [...p, ...add];
    });
    setCerts(p => {
      const names = p.map(x => x.name);
      return [...p, ...result.certs.filter(c => c._include && !names.includes(c.name)).map(c => ({ id: "c_" + Date.now() + Math.random().toString(36).slice(2, 4), name: c.name, issuer: c.issuer, date: c.date, note: c.note }))];
    });
    if (result.profile && result._profileInclude) {
      setResumeProfile(p => ({ ...p, ...Object.fromEntries(Object.entries(result.profile).filter(([, v]) => v)), _imported: true }));
    }
    setPhase("done");
  };

  if (phase === "input" || phase === "loading") return (
    <div style={{ maxWidth: 680 }}>
      <H2>자료 넣기 / 경험 정리하기</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.65 }}>
        기존 이력서·경험 정리 파일이 정형화되어 있지 않아도 괜찮습니다. 파일이 없다면 그냥 아래에 자유롭게 적어도 됩니다.<br />
        AI가 내용을 추출해 초안을 만들면, <b>모든 항목을 직접 확인·수정한 뒤</b> 반영합니다. 확인 전에는 아무것도 저장되지 않습니다.
      </div>
      <Card>
        <Label>파일 선택 (.txt / .md / .docx / .xlsx / .xls / .csv) 또는 내용 직접 작성</Label>
        <input type="file" accept=".txt,.md,.docx,.xlsx,.xls,.csv" onChange={onFile} style={{ fontFamily: font, fontSize: "var(--fs-base)", marginBottom: 10 }} />
        {fileName && <div style={{ fontSize: "var(--fs-sm)", color: C.blueText, marginBottom: 8 }}>선택됨: {fileName}</div>}

        {workbook && (
          <div style={{ marginBottom: 10, padding: "var(--sp-5) var(--sp-5)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)" }}>
            <div style={{ fontSize: "var(--fs-base)", fontWeight: 700, marginBottom: 8 }}>시트가 여러 개 있습니다 — 가져올 시트를 선택하세요</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {workbook.SheetNames.map(name => (
                <Btn key={name} small onClick={() => loadSheet(workbook, name)}>{name}</Btn>
              ))}
            </div>
          </div>
        )}

        <Textarea rows={11} placeholder="이미 작성된 자소서 등의 파일이 있다면 첨부해주세요! 저희가 정리해드립니다. (파일 없이 여기에 바로 적어도 괜찮아요)" value={raw} onChange={e => setRaw(e.target.value)} />
        {raw.length > 12000 && (
          <div style={{ marginTop: 8, fontSize: "var(--fs-sm)", color: C.sub }}>
            원문이 {raw.length.toLocaleString()}자로 깁니다. 너무 길면 API 오류가 날 수 있으니, 관련 없는 시트·행은 미리 지우고 필요한 부분만 남기는 걸 권장합니다.
          </div>
        )}
        {error && (
          <div style={{ marginTop: 10, padding: "var(--sp-4) var(--sp-5)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)" }}>
            <div style={{ fontSize: "var(--fs-sm)", fontWeight: 700, color: C.redText, marginBottom: 4 }}>오류</div>
            <div style={{ fontSize: "var(--fs-sm)", color: C.redText, whiteSpace: "pre-wrap", fontFamily: "monospace", lineHeight: 1.5 }}>{error}</div>
          </div>
        )}
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}>
          <Btn primary disabled={!raw.trim() || phase === "loading"} onClick={extract}>
            {phase === "loading" ? "AI 추출 중…" : "AI로 추출하기 →"}
          </Btn>
        </div>
      </Card>
      <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginTop: 12, lineHeight: 1.6 }}>
        AI 추출 규칙: 원문에 없는 내용은 만들지 않으며, 모든 수치는 근거 자료가 확인될 때까지 "추가 확인 필요" 상태로 들어옵니다.<br />
        입력한 내용은 추출을 위해 외부 AI 서버로 전송되며, 이 앱이 별도로 저장하지 않습니다.
      </div>
    </div>
  );

  if (phase === "review") return (
    <div style={{ maxWidth: 780 }}>
      <H2>추출 결과 검토 — 반영 전 확인·수정</H2>
      <div style={{ display: "flex", gap: 8, padding: "var(--sp-4) var(--sp-5)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)", marginBottom: 16, alignItems: "center" }}>
        <Badge label="AI 추출" color={C.ai} bg="#fff" />
        <span style={{ fontSize: "var(--fs-base)", color: C.text }}>모든 필드를 수정할 수 있습니다. 체크 해제한 항목은 반영되지 않습니다.</span>
      </div>

      {(result.profile.name || result.profile.email) && (
        <>
          <Label>기본 정보 → 기본 이력서</Label>
          <Card style={{ marginBottom: 16 }}>
            <label style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, fontSize: "var(--fs-base)", cursor: "pointer" }}>
              <input type="checkbox" checked={result._profileInclude} onChange={() => setResult(p => ({ ...p, _profileInclude: !p._profileInclude }))} />
              기본 이력서에 반영
            </label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, opacity: result._profileInclude ? 1 : 0.45 }} className="stack-sm">
              <div><Label>이름</Label><Input value={result.profile.name || ""} onChange={e => patchProfile("name", e.target.value)} /></div>
              <div><Label>이메일</Label><Input value={result.profile.email || ""} onChange={e => patchProfile("email", e.target.value)} /></div>
              {result.profile.targetRole && <div style={{ gridColumn: "1 / -1" }}><Label>목표 직무</Label><Input value={result.profile.targetRole || ""} onChange={e => patchProfile("targetRole", e.target.value)} /></div>}
            </div>
          </Card>
        </>
      )}

      <Label>경험 → 경험 보관함 (초기 메모 상태로 추가)</Label>
      {result.experiences.map(e => (
        <Card key={e._id} style={{ marginBottom: 10, opacity: e._include ? 1 : 0.45 }}>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
            <input type="checkbox" checked={e._include} onChange={() => toggle("experiences", e._id)} style={{ marginTop: 8 }} />
            <div style={{ flex: 1, display: "grid", gap: 8 }}>
              {e._duplicate && (
                <div style={{ fontSize: "var(--fs-sm)", color: C.orangeText, background: C.accent, border: `1px solid ${C.line}`, padding: "var(--sp-2) var(--sp-4)", borderRadius: "var(--r-lg)" }}>
                  이미 보관함에 같은 제목·소속의 경험이 있습니다. 중복일 가능성이 있어 기본적으로 체크가 해제되어 있습니다.
                </div>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1.3fr 1.3fr", gap: 8 }} className="tbl-row">
                <Input value={e.title} onChange={ev => patchExp(e._id, "title", ev.target.value)} />
                <Input value={e.organization || ""} placeholder="소속" onChange={ev => patchExp(e._id, "organization", ev.target.value)} />
                <Input value={e.role || ""} placeholder="역할" onChange={ev => patchExp(e._id, "role", ev.target.value)} />
              </div>
              <Textarea style={{ minHeight: 54 }} value={e.rawNote || ""} onChange={ev => patchExp(e._id, "rawNote", ev.target.value)} />
              {(e.competencies || []).length > 0 && (
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                  {e.competencies.map((c, i) => <Badge key={i} label={"#" + c} color={C.blueText} bg={C.blueBg} />)}
                </div>
              )}
              {(e.metrics || []).map((m, mi) => (
                <div key={mi} style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", background: C.accent, border: `1px solid ${C.line}`, padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)" }} className="wrap-sm">
                  <Badge label="추가 확인 필요" color={C.redText} bg={C.redBg} />
                  <Input value={m.metricName} onChange={ev => patchMetric(e._id, mi, "metricName", ev.target.value)} style={{ width: 120 }} />
                  {m.beforeValue != null ? (
                    <>
                      <Input value={m.beforeValue} onChange={ev => patchMetric(e._id, mi, "beforeValue", ev.target.value)} style={{ width: 70, textAlign: "right" }} />
                      <span style={{ color: C.faintText }}>→</span>
                      <Input value={m.afterValue} onChange={ev => patchMetric(e._id, mi, "afterValue", ev.target.value)} style={{ width: 70 }} />
                    </>
                  ) : (
                    <Input value={m.changeValue} onChange={ev => patchMetric(e._id, mi, "changeValue", ev.target.value)} style={{ width: 70 }} />
                  )}
                  <Input value={m.unit || ""} onChange={ev => patchMetric(e._id, mi, "unit", ev.target.value)} style={{ width: 50 }} />
                  <span style={{ fontSize: "var(--fs-xs)", color: C.sub, flex: "1 1 100%" }}>{m.note}</span>
                </div>
              ))}
              <div style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>수치는 경험 분석의 성과 단계에서 근거를 확인해야 정식 수치(단일 원본)로 승격됩니다. 역량 태그는 AI 추정이니 경험 상세에서 다시 확인하세요.</div>
            </div>
          </div>
        </Card>
      ))}

      <Label>스킬 → 역량·스킬</Label>
      <Card style={{ marginBottom: 10 }}>
        {result.skills.map(s => (
          <div key={s._id} style={{ padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, opacity: s._include ? 1 : 0.45 }}>
            <label style={{ display: "flex", gap: 10, alignItems: "center", cursor: "pointer", marginBottom: 6 }}>
              <input type="checkbox" checked={s._include} onChange={() => toggle("skills", s._id)} />
              <span style={{ fontWeight: 600, fontSize: "var(--fs-base)" }}>{s.name}</span>
            </label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", paddingLeft: 26 }}>
              {(s.scopeItems || []).map((it, i) => <Badge key={i} label={it.text} color={C.sub} bg={C.lineSoft} />)}
              <Badge label="활용 범위 확인 필요" color={C.orangeText} bg={C.orangeBg} />
            </div>
          </div>
        ))}
      </Card>

      <Label>자격증 → 역량·스킬 / 기본 이력서</Label>
      <Card style={{ marginBottom: 16 }}>
        {result.certs.map(c => (
          <label key={c._id} style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", padding: "var(--sp-2) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)", cursor: "pointer", opacity: c._include ? 1 : 0.45 }}>
            <input type="checkbox" checked={c._include} onChange={() => toggle("certs", c._id)} />
            <span style={{ fontWeight: 600, flexShrink: 0 }}>{c.name}</span>
            <span style={{ color: C.sub, fontSize: "var(--fs-sm)" }}>{c.date || "취득일 미상"}</span>
            {c.note && <span style={{ color: C.orangeText, fontSize: "var(--fs-sm)" }}>{c.note}</span>}
          </label>
        ))}
      </Card>

      <div style={{ display: "flex", justifyContent: "space-between" }} className="wrap-sm">
        <Btn onClick={() => setPhase("input")}>← 다시 추출</Btn>
        <Btn primary onClick={commit}>확인한 항목 반영하기</Btn>
      </div>
    </div>
  );

  return (
    <div style={{ maxWidth: 640 }}>
      <H2>반영 완료</H2>
      <Card>
        <div style={{ fontSize: "var(--fs-md)", lineHeight: 1.7 }}>
          선택한 항목이 <b>초기 메모</b> 상태로 추가되었습니다.<br />
          경험 보관함에서 각 경험의 <b>단계별 분석</b>을 진행하면, 가져온 수치도 근거 확인 후 정식 수치로 승격됩니다.
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 14 }} className="wrap-sm">
          <Btn primary onClick={() => { const last = experiences[experiences.length - 1]; last && onDone(last.id); }}>경험 보관함 보기 →</Btn>
          <Btn onClick={() => { setPhase("input"); setRaw(""); setResult(null); setFileName(""); }}>다른 파일 가져오기</Btn>
        </div>
      </Card>
    </div>
  );
}

/* ============================================================ 역량·스킬 */
function Skills({ skills, setSkills, experiences, onOpenExp, addTrash }) {
  const [tab, setTab] = useState("도구");
  const [newSkill, setNewSkill] = useState("");
  const [addingScope, setAddingScope] = useState(null); // skillId
  const [scopeDraft, setScopeDraft] = useState({ text: "", evidenceExpId: "" });

  const byCat = (cat) => skills.filter(s => s.category === cat);
  const catOf = { "도구": "tool", "직무 역량": "skill" };

  const addSkill = (cat) => {
    if (!newSkill.trim()) return;
    setSkills(p => [...p, { id: "s_" + Date.now(), name: newSkill.trim(), category: cat, summary: "", scopeItems: [] }]);
    setNewSkill("");
  };
  const addScope = (skillId) => {
    if (!scopeDraft.text.trim()) return;
    setSkills(p => p.map(s => s.id === skillId ? { ...s, scopeItems: [...s.scopeItems, { id: "sc_" + Date.now(), text: scopeDraft.text.trim(), evidenceExpId: scopeDraft.evidenceExpId || null }] } : s));
    setScopeDraft({ text: "", evidenceExpId: "" }); setAddingScope(null);
  };
  const removeScope = (skillId, scId) =>
    setSkills(p => p.map(s => s.id === skillId ? { ...s, scopeItems: s.scopeItems.filter(x => x.id !== scId) } : s));
  const patchSkill = (skillId, k, v) => setSkills(p => p.map(s => s.id === skillId ? { ...s, [k]: v } : s));
  const removeSkill = (skillId) => {
    const skill = skills.find(s => s.id === skillId);
    setSkills(p => p.filter(s => s.id !== skillId));
    addTrash("skill", skill.name, skill);
  };

  const SkillCard = ({ s }) => {
    const linked = s.scopeItems.filter(i => i.evidenceExpId).length;
    return (
      <Card style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4, gap: 10 }}>
          <Input value={s.name} onChange={e => patchSkill(s.id, "name", e.target.value)}
            style={{ fontWeight: 700, fontSize: "var(--fs-md)", border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
          <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
            <span style={{ fontSize: "var(--fs-xs)", color: C.faintText, whiteSpace: "nowrap" }}>근거 연결 {linked}/{s.scopeItems.length}</span>
            <span {...clickableProps(() => removeSkill(s.id), { label: "닫기" })} title="이 항목 삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
          </div>
        </div>
        <Input value={s.summary || ""} placeholder="한 줄 요약 (선택)" onChange={e => patchSkill(s.id, "summary", e.target.value)}
          style={{ fontSize: "var(--fs-sm)", color: C.sub, border: "none", padding: "var(--sp-1) 0", marginBottom: 10 }} />

        <Label>활용 범위 — 할 수 있는 작업을 구체적으로</Label>
        {s.scopeItems.map(item => {
          const exp = experiences.find(e => e.id === item.evidenceExpId);
          return (
            <div key={item.id} style={{ display: "flex", gap: 8, alignItems: "center", padding: "var(--sp-2) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }}>
              <span style={{ flex: 1, lineHeight: 1.5 }}>{item.text}</span>
              {exp ? (
                <span {...clickableProps(() => onOpenExp(exp.id))} style={{ fontSize: "var(--fs-xs)", color: C.blueText, background: "transparent", border: `1px solid ${C.blue}55`, padding: "var(--sp-1) var(--sp-2)", borderRadius: "var(--r-lg)", cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" }}>
                  {exp.title} →
                </span>
              ) : (
                <Badge label="경험 근거 없음" color={C.orangeText} bg={C.orangeBg} />
              )}
              <span {...clickableProps(() => removeScope(s.id, item.id), { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
            </div>
          );
        })}

        {addingScope === s.id ? (
          <div style={{ display: "flex", gap: 8, marginTop: 10 }} className="wrap-sm">
            <Input autoFocus placeholder="예: 피벗 테이블 기반 판매 데이터 집계" value={scopeDraft.text}
              onChange={e => setScopeDraft(d => ({ ...d, text: e.target.value }))} onKeyDown={e => e.key === "Enter" && addScope(s.id)} style={{ flex: 1 }} />
            <select value={scopeDraft.evidenceExpId} onChange={e => setScopeDraft(d => ({ ...d, evidenceExpId: e.target.value }))}
              style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-3) var(--sp-4)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, background: C.panel, maxWidth: 200 }}>
              <option value="">근거 경험 (선택)</option>
              {experiences.map(e => <option key={e.id} value={e.id}>{e.title}</option>)}
            </select>
            <Btn small primary onClick={() => addScope(s.id)}>추가</Btn>
            <Btn small onClick={() => setAddingScope(null)}>취소</Btn>
          </div>
        ) : (
          <div style={{ marginTop: 10 }}><Btn small onClick={() => { setAddingScope(s.id); setScopeDraft({ text: "", evidenceExpId: "" }); }}>+ 활용 범위 추가</Btn></div>
        )}
      </Card>
    );
  };

  return (
    <div style={{ maxWidth: 760 }}>
      <H2>역량·스킬</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 6, lineHeight: 1.6 }}>
        "상·중·하" 자기 평가 대신 <b>실제로 할 수 있는 작업</b>을 적고, 경험 근거를 연결합니다.
      </div>
      <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginBottom: 16 }}>
        근거가 연결된 항목만 이력서·자소서에서 자신 있게 쓸 수 있습니다. 근거 없는 항목은 면접 검증 리스크가 있습니다.
      </div>

      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.line}`, marginBottom: 18 }}>
        {["도구", "직무 역량"].map(t => (
          <div key={t} {...clickableProps(() => setTab(t))} style={{ padding: "var(--sp-3) var(--sp-5)", fontSize: "var(--fs-base)", fontWeight: tab === t ? 700 : 500, cursor: "pointer",
            color: tab === t ? C.text : C.sub, borderBottom: tab === t ? `2px solid ${C.text}` : "2px solid transparent", marginBottom: -1 }}>{t}</div>
        ))}
      </div>

      {(tab === "도구" || tab === "직무 역량") && (
        <>
          {byCat(catOf[tab]).map(s => <SkillCard key={s.id} s={s} />)}
          <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
            <Input placeholder={tab === "도구" ? "도구 이름 (예: Google Analytics)" : "역량 이름 (예: 상품 기획)"} value={newSkill}
              onChange={e => setNewSkill(e.target.value)} onKeyDown={e => e.key === "Enter" && addSkill(catOf[tab])} style={{ maxWidth: 320 }} />
            <Btn onClick={() => addSkill(catOf[tab])}>+ 추가</Btn>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================================================ 지원 관리 */
function Applications({ applications, setApplications, onOpen, addTrash }) {
  const addApp = () => {
    const id = "ap_" + Date.now();
    setApplications(prev => [...prev, { id, company: "새 지원처", position: "", deadline: "", status: "interested", priority: "medium",
      essayProgress: 0, interviewProgress: 0, requirements: [], essays: [], interviews: [] }]);
    onOpen(id);
  };
  const deleteApp = (id) => {
    const app = applications.find(a => a.id === id);
    setApplications(prev => prev.filter(a => a.id !== id));
    addTrash("application", `${app.company} ${app.position}`.trim(), app);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <H2>지원 관리</H2>
        <Btn primary onClick={addApp}>+ 지원 등록</Btn>
      </div>
      {applications.map(a => (
        <Card key={a.id} onClick={() => onOpen(a.id)} style={{ marginBottom: 10, display: "grid", gridTemplateColumns: "1.5fr 1fr 100px 130px 130px 20px", alignItems: "center", gap: 10 }} className="tbl-row">
          <div><span style={{ fontWeight: 700, fontSize: "var(--fs-md)" }}>{a.company}</span><span style={{ color: C.sub, fontSize: "var(--fs-base)", marginLeft: 8 }}>{a.position}</span></div>
          <div style={{ fontSize: "var(--fs-base)", color: C.sub }}>마감 {a.deadline || "미정"}</div>
          <Badge label={{ interested: "관심", analyzing: "분석 중", writing: "작성 중", submitted: "제출", interview: "면접", result: "결과" }[a.status]} color={C.blueText} bg={C.blueBg} />
          <div style={{ fontSize: "var(--fs-sm)", color: C.sub }}>자소서 {a.essayProgress}%</div>
          <div style={{ fontSize: "var(--fs-sm)", color: C.sub }}>면접 준비 {a.interviewProgress}%</div>
          <span {...clickableProps(ev => { ev.stopPropagation(); deleteApp(a.id); }, { label: "닫기" })}
            title="삭제 (휴지통에서 복구 가능)" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
        </Card>
      ))}
      {applications.length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.faintText }}>등록된 지원처가 없습니다. "+ 지원 등록"으로 추가하세요.</div>}
    </div>
  );
}

/* ============================================================ 자소서 작성 챗봇 */
const ESSAY_COACH_SYSTEM_PROMPT = `지금부터 당신은 국내 대기업·외국계·스타트업 채용을 모두 경험한 시니어 채용담당자이자, 수천 건 이상의 합격 자기소개서를 첨삭한 커리어 컨설턴트입니다.
내가 아래와 같은 정보를 순서와 형식에 관계없이 제공할 것입니다.
   * 내 경력 및 경험
   * 이력서(Resume/CV)
   * 지원하려는 회사와 직무(Job Description)
   * 자기소개서 문항(있는 경우)
   * 추가로 강조하고 싶은 내용이나 피하고 싶은 표현
당신의 역할은 단순히 글을 작성하는 것이 아니라, 지원자의 경험을 채용담당자의 시각에서 가장 설득력 있게 재구성하는 것입니다.

반드시 수행해야 하는 작업
먼저 내가 제공한 정보를 분석하여 다음을 수행하십시오.
   1. 지원 직무에서 가장 중요하게 평가할 역량을 추론합니다.
   2. 내 이력서와 경험 중 어떤 사례가 가장 설득력이 높은지 선별합니다.
   3. 부족한 정보가 있다면 자소서를 쓰기 전에 반드시 질문합니다.
   4. 경험이 여러 개라면 가장 경쟁력 있는 스토리를 우선 추천하고, 그 이유도 간단히 설명합니다.
충분한 정보가 확보되면 자기소개서 작성을 시작하십시오.

작성 원칙
   * 절대 경험을 과장하거나 허위 사실을 만들어내지 마십시오.
   * 아래 제공되는 "사실 정보"에 없는 내용은 지어내지 마십시오. 사실 정보에 없는 수치나 성과는 절대 임의로 만들지 마십시오.
   * 추상적인 표현보다 실제 행동과 과정(Action)을 중심으로 작성하십시오.
   * "책임감이 강합니다", "열심히 했습니다"와 같은 진부한 표현은 사용하지 마십시오.
   * 내가 실제 수행한 업무, 사용한 도구, 의사결정 과정, 문제 해결 방식이 드러나도록 작성하십시오.
   * STAR, CAR 등의 구조를 참고하되 자연스럽게 녹여내십시오.
   * 결과보다 왜 그런 판단을 했는지, 어떻게 해결했는지가 드러나는 글을 작성하십시오.
   * 채용담당자가 읽기 쉬운 두괄식 구조를 유지하십시오.
   * 문항별 글자 수 제한이 있다면 90~95% 수준까지 작성하고, 제한이 없다면 공백 포함 약 700자 내외를 기본으로 합니다.
   * 전문적이고 담백한 경어체를 유지하십시오.

작성 방식
한 번에 모든 문항을 작성하지 마십시오.
반드시 다음 순서를 따르십시오.
   1. 먼저 어떤 경험을 사용할 것인지 추천합니다.
   2. 그 경험을 사용하는 이유를 설명합니다.
   3. 1번 문항만 작성합니다.
   4. 이후 사용자의 피드백을 기다립니다.
   5. 수정 요청이 있으면 즉시 반영합니다.
   6. 사용자가 "다음 문항"이라고 말하면 다음 문항을 작성합니다.

문항 작성 형식
각 문항은 다음 형식을 따르십시오.
   * 핵심 메시지를 담은 소제목 1개
   * 두괄식 첫 문장
   * 행동(Action) 중심의 본문
   * 결과와 직무 적합성으로 마무리

피드백 원칙
초안을 작성한 뒤에는 다음을 함께 제공하십시오.
   * 채용담당자 관점에서 가장 강한 부분
   * 더 보완하면 좋은 부분
   * 더 설득력 있게 만들기 위해 필요한 추가 정보(있다면)

작성이 끝나면 다음 안내만 덧붙이십시오.
"초안을 검토해 보시고 수정하고 싶은 부분(분량, 강조점, 표현 등)을 말씀해 주세요. 마음에 드신다면 '다음 문항'이라고 입력해 주세요."`;

const PERSONAL_ASSISTANT_SYSTEM_PROMPT = `당신은 사용자의 취업 준비를 옆에서 도와주는 친근한 개인 어시스턴트입니다. 채용담당자나 컨설턴트 페르소나가 아니라, 사용자의 경험과 상황을 잘 아는 친구 같은 존재입니다.

역할
사용자가 취업 준비 중 드는 개인적인 고민이나 사소한 질문(예: "이 회사 지원할까 말까", "내 경험 중에 뭐가 제일 강점인 것 같아?", "요즘 너무 불안한데 어떻게 해야 할까", "이 자격증 딸 가치가 있을까")에, 아래 제공되는 사용자의 실제 데이터(경험/역량/자격증/지원 현황)를 참고해서 답합니다.

원칙
- 아래 데이터에 없는 사실을 지어내지 마십시오. 데이터에 없으면 "그 부분은 아직 정리가 안 되어 있네요"라고 솔직히 말하십시오.
- 채용담당자처럼 평가하거나 심사하는 톤을 쓰지 마십시오. 옆에서 같이 고민해주는 톤을 쓰십시오.
- 진로·심리적으로 무거운 고민이면 성급하게 정답을 주기보다 사용자의 상황을 먼저 이해하려는 질문을 해도 됩니다.
- 사소한 질문(맞춤법, 이 표현이 나은지 등)은 바로 간단히 답하십시오.
- 답변은 짧고 자연스럽게. 보고서처럼 항목별로 나열하지 말고, 대화하듯 쓰십시오.`;

function buildPersonalContext(experiences, skills, certs, awards, resumeProfile, applications, metrics) {
  const expLines = experiences.map(e => {
    const myMetrics = (metrics || []).filter(m => m.experienceId === e.id);
    const metricStr = myMetrics.length ? ` [수치: ${myMetrics.map(m => `${m.metricName} ${formatMetric(m, "exact")}`).join(", ")}]` : "";
    const extras = [
      e.goal ? `목표: ${e.goal}` : null,
      e.difficulty ? `어려움: ${e.difficulty}` : null,
      e.learning ? `배운 점: ${e.learning}` : null,
      e.jobRelevance ? `직무 연결: ${e.jobRelevance}` : null,
    ].filter(Boolean).join(" / ");
    return `- ${e.title} (${e.organization || "소속 미상"}, ${e.status}) — ${e.oneLineSummary || e.context || "요약 없음"}${metricStr}${(e.competencies || []).length ? ` [역량: ${e.competencies.join(", ")}]` : ""}${extras ? ` [${extras}]` : ""}`;
  }).join("\n");
  const skillLines = (skills || []).map(s => `- ${s.name}`).join(", ");
  const certLines = (certs || []).map(c => `- ${c.name}${c.date ? ` (${c.date})` : ""}`).join(", ");
  const awardLines = (awards || []).map(a => `- ${a.name}`).join(", ");
  const appLines = (applications || []).map(a => `- ${a.company} · ${a.position} (${a.status}${a.deadline ? `, 마감 ${a.deadline}` : ""})`).join("\n");

  return `[사용자 프로필]
이름: ${resumeProfile?.name || "미입력"} / 희망 직무: ${resumeProfile?.targetRole || "미입력"}
한 줄 소개: ${resumeProfile?.headline || "미입력"}

[정리된 경험 — ${experiences.length}건]
${expLines || "아직 정리된 경험이 없습니다."}

[역량·스킬]
${skillLines || "없음"}

[자격증·어학]
${certLines || "없음"}

[수상기록]
${awardLines || "없음"}

[지원 현황]
${appLines || "등록된 지원처가 없습니다."}`;
}

const EXPERIENCE_REVIEW_SYSTEM_PROMPT = `당신은 국내 대기업·외국계·스타트업 채용을 두루 경험한 시니어 채용담당자입니다. 지금부터 지원자가 정리한 "경험 데이터베이스" 전체를 검토합니다.

역할
지원자가 이 데이터를 이력서·자소서·면접에 그대로 활용할 것이므로, 실무자 시선에서 부족한 부분을 냉정하게 짚어주는 것이 당신의 역할입니다. 무조건적인 칭찬은 도움이 되지 않습니다.

검토 시 반드시 확인할 것
   * 수치·성과가 빠져 있거나 모호한 경험 (예: "매출이 늘었다" 수준에서 멈춘 경우)
   * 배경·문제·행동·기여도·성과 중 비어 있는 항목이 있는 경험
   * 본인이 한 일과 팀이 한 일이 구분되지 않는 경험 (기여도 근거 부실)
   * 서로 내용이 겹치거나 같은 일화가 여러 경험으로 쪼개져 있는 것으로 보이는 경우 (합치기를 제안할 것)
   * 어려움·배운 점이 비어 있어 면접 압박 질문(실패, 갈등, 어려움 극복)에 쓸 수 없는 경험
   * 전체적으로 부족한 역량 유형 (예: 리더십, 협업, 문제 해결 중 특정 유형의 경험이 없는 경우)

답변 방식
   * 한 번에 모든 걸 나열하지 말고, 가장 시급하고 임팩트가 큰 문제 3~5가지를 우선순위대로 짚으십시오.
   * 각 지적에는 어떤 경험(제목)의 어떤 부분이 문제인지 구체적으로 명시하십시오.
   * 추상적인 조언("더 구체적으로 쓰세요") 대신, 무엇을 확인하거나 채워 넣으면 되는지 실행 가능한 다음 행동을 제시하십시오.
   * 절대 지어내지 마십시오 — 데이터에 없는 내용을 추측해서 "이랬을 것이다"라고 단정하지 말고, 없으면 "확인이 필요합니다"라고 하십시오.
   * 사용자가 특정 경험에 대해 더 파고들어 질문하면 그 경험에 집중해서 답하십시오.`;

function buildReviewContext(experiences, metrics) {
  const lines = experiences.map(e => {
    const parts = [`- [${e.title}] (${e.organization || "소속 미상"} · ${e.status}${e.depthDone ? "" : " · 심화 미입력"})`];
    parts.push(`  배경: ${e.context || "(없음)"}`);
    parts.push(`  문제: ${e.discoveredProblem || "(없음)"}`);
    parts.push(`  본인 기여: ${e.personalContribution || "(없음)"} / 기여 근거: ${e.contributionEvidence || "(없음)"}`);
    if (e.goal) parts.push(`  목표: ${e.goal}`);
    if (e.actions?.length) {
      parts.push(`  행동:\n${e.actions.map(a => `    · (${ACTION_LABEL[a.actionType] || a.actionType}) ${a.description}`).join("\n")}`);
    }
    const myMetrics = (metrics || []).filter(m => m.experienceId === e.id);
    if (myMetrics.length) {
      parts.push(`  성과 수치:\n${myMetrics.map(m => `    · ${m.metricName}: ${formatMetric(m, "exact")} (${CERTAINTY[m.certainty]?.[0] || m.certainty})`).join("\n")}`);
    }
    parts.push(`  성과 요약: ${e.oneLineSummary || "(없음)"}`);
    if (e.qualitative) parts.push(`  정성 성과: ${e.qualitative}`);
    parts.push(`  어려움: ${e.difficulty || "(없음)"} / 배운 점: ${e.learning || "(없음)"}`);
    if (e.coreMessage) parts.push(`  핵심 메시지: ${e.coreMessage}`);
    if (e.jobRelevance) parts.push(`  직무 연결: ${e.jobRelevance}`);
    parts.push(`  역량 태그: ${(e.competencies || []).join(", ") || "(없음)"}`);
    return parts.join("\n");
  }).join("\n\n");

  return `[지원자의 경험 데이터베이스 전체 — ${experiences.length}건]
${lines || "등록된 경험이 없습니다."}`;
}

function buildEssayContext(app, essay, experiences, metrics) {
  const describeExp = (e) => {
    const parts = [`- [${e.title}] ${e.organization || ""} · ${e.role || ""} (${e.startDate}~${e.endDate})`];
    if (e.context) parts.push(`  배경: ${e.context}`);
    if (e.discoveredProblem) parts.push(`  문제: ${e.discoveredProblem}`);
    if (e.personalContribution) parts.push(`  본인 행동/기여: ${e.personalContribution}`);
    if (e.contributionEvidence) parts.push(`  기여 근거: ${e.contributionEvidence}`);
    if (e.goal) parts.push(`  목표: ${e.goal}`);
    if (e.actions?.length) {
      const actionLines = e.actions.map(a => `    · (${ACTION_LABEL[a.actionType] || a.actionType}) ${a.description}`).join("\n");
      parts.push(`  행동 타임라인:\n${actionLines}`);
    }
    const myMetrics = (metrics || []).filter(m => m.experienceId === e.id);
    if (myMetrics.length) {
      const metricLines = myMetrics.map(m =>
        `    · ${m.metricName}: ${formatMetric(m, "exact")} (${CERTAINTY[m.certainty]?.[0] || m.certainty}${m.evidenceSource ? `, 근거: ${m.evidenceSource}` : ""})`
      ).join("\n");
      parts.push(`  성과 수치 (certainty가 "확인 필요"인 값은 문장에 그대로 확정적으로 쓰지 말고 사용자에게 확인을 권할 것):\n${metricLines}`);
    }
    if (e.oneLineSummary) parts.push(`  성과 요약: ${e.oneLineSummary}`);
    if (e.qualitative) parts.push(`  정성 성과: ${e.qualitative}`);
    if (e.difficulty) parts.push(`  어려움: ${e.difficulty}`);
    if (e.learning) parts.push(`  배운 점: ${e.learning}`);
    if (e.coreMessage) parts.push(`  핵심 메시지: ${e.coreMessage}`);
    if (e.jobRelevance) parts.push(`  직무 연결: ${e.jobRelevance}`);
    if (e.competencies?.length) parts.push(`  관련 역량: ${e.competencies.join(", ")}`);
    return parts.join("\n");
  };

  const selected = (essay.selectedExperienceIds || []).map(id => experiences.find(e => e.id === id)).filter(Boolean);
  const usable = selected.length > 0 ? selected : experiences.filter(e => e.status !== "draft");
  const others = selected.length > 0
    ? experiences.filter(e => e.status !== "draft" && !selected.some(s => s.id === e.id))
    : [];

  const factLines = usable.map(describeExp).join("\n\n");
  const otherTitles = others.map(e => `- ${e.title}`).join("\n");

  const reqLines = (app.requirements || []).map(r => `- ${r.requirement} (중요도 ${r.importance}/5)${r.matchReason ? ` — ${r.matchReason}` : ""}`).join("\n");

  return `[지원 정보]
회사: ${app.company}
직무: ${app.position}

[이 회사가 요구하는 역량 (공고 분석 결과) — 답변에서 이 키워드와 최대한 연결지어 서술할 것]
${reqLines || "등록된 요구 역량 없음"}

[자기소개서 문항]
"${essay.question || "(문항 미입력 — 사용자에게 문항을 먼저 물어볼 것)"}"
글자 수 제한: ${essay.characterLimit}자

[이 문항에 사용하기로 선택된 경험 — 우선적으로 이것만 활용하고, 여기 없는 내용은 지어내지 말 것]
${factLines || "선택되었거나 분석 완료된 경험이 없습니다. 먼저 어떤 경험을 쓸지 사용자에게 추천/확인하십시오."}
${others.length > 0 ? `\n[그 외 참고 가능한 경험 (제목만) — 선택된 경험이 문항과 잘 안 맞아 보이면 이 중에서 대안을 제안할 것]\n${otherTitles}` : ""}`;
}

function EssayChat({ title, subtitle, systemPrompt, contextText, autoStartMessage, inputPlaceholder, onClose, closeLabel, onSaveDraft, saveDraftLabel, history, onHistoryChange }) {
  const [messages, setMessages] = useState(history || []); // {role, content}
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const bottomRef = useRef(null);
  const started = useRef(false);

  const updateMessages = (updater) => {
    setMessages(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      onHistoryChange && onHistoryChange(next);
      return next;
    });
  };

  useEffect(() => {
    if (!started.current) {
      started.current = true;
      if ((history || []).length === 0) {
        send(autoStartMessage || "안녕하세요, 도와주세요.", true);
      }
    }
    // eslint-disable-next-line
  }, []);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages, loading]);

  const send = async (text, isAutoStart = false) => {
    if (!text.trim() && !isAutoStart) return;
    const nextMessages = [...messages, { role: "user", content: text }];
    updateMessages(nextMessages);
    setInput(""); setLoading(true); setError("");
    try {
      const res = await fetch("/api/chat", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          systemPrompt,
          context: contextText,
          messages: nextMessages,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data?.error?.message
          || (typeof data?.error === "string" ? data.error : null)
          || `API 오류 (HTTP ${res.status}) — 응답 원문: ${JSON.stringify(data).slice(0, 500)}`);
      }
      const reply = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("\n");
      if (!reply) throw new Error("응답에 텍스트가 없습니다.");
      updateMessages(prev => [...prev, { role: "assistant", content: reply }]);
    } catch (e) {
      setError(e.message || String(e));
      updateMessages(prev => prev.slice(0, isAutoStart ? 0 : -1));
    } finally {
      setLoading(false);
    }
  };

  const lastAssistant = [...messages].reverse().find(m => m.role === "assistant");

  return (
    <Card style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "var(--sp-5) var(--sp-6)", borderBottom: `1px solid ${C.line}` }} className="wrap-sm">
        <div>
          <div style={{ fontSize: "var(--fs-base)", fontWeight: 700 }}>{title}</div>
          {subtitle && <div style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>{subtitle}</div>}
          <div style={{ fontSize: "var(--fs-2xs)", color: C.faintText, marginTop: 2 }}>이 대화 내용은 응답 생성을 위해 외부 AI 서버로 전송됩니다</div>
        </div>
        <div style={{ display: "flex", gap: 8 }} className="wrap-sm">
          {messages.length > 0 && <Btn small onClick={() => { updateMessages([]); started.current = false; }}>대화 초기화</Btn>}
          {onSaveDraft && lastAssistant && <Btn small onClick={() => onSaveDraft(lastAssistant.content)}>{saveDraftLabel || "이 답변을 초안으로 저장"}</Btn>}
          <Btn small onClick={onClose}>{closeLabel || "← 목록으로"}</Btn>
        </div>
      </div>

      <div style={{ height: 420, overflowY: "auto", padding: "var(--sp-6)", background: C.bg }}>
        {messages.filter(m => m.role !== "system").map((m, i) => (
          <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", marginBottom: 10 }}>
            <div style={{
              maxWidth: "80%", padding: "var(--sp-4) var(--sp-5)", borderRadius: "var(--r-lg)", fontSize: "var(--fs-base)", lineHeight: 1.6, whiteSpace: "pre-wrap",
              background: m.role === "user" ? C.text : C.panel, color: m.role === "user" ? "#fff" : C.text,
              border: m.role === "user" ? "none" : `1px solid ${C.line}` }}>
              {m.content}
            </div>
          </div>
        ))}
        {loading && <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>답변을 작성하는 중…</div>}
        {error && (
          <div style={{ padding: "var(--sp-4) var(--sp-5)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-lg)" }}>
            <div style={{ fontSize: "var(--fs-sm)", fontWeight: 700, color: C.redText, marginBottom: 4 }}>오류</div>
            <div style={{ fontSize: "var(--fs-sm)", color: C.redText, whiteSpace: "pre-wrap", fontFamily: "monospace" }}>{error}</div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      <div style={{ display: "flex", gap: 8, padding: "var(--sp-5)", borderTop: `1px solid ${C.line}` }} className="wrap-sm">
        <Textarea rows={2} placeholder={inputPlaceholder || "피드백을 입력하거나 '다음 문항'이라고 입력하세요"} value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(input); } }}
          style={{ flex: 1, minHeight: 44 }} />
        <Btn primary disabled={loading || !input.trim()} onClick={() => send(input)}>전송</Btn>
      </div>
    </Card>
  );
}

function JobPostingExtractor({ experiences, raw, onExtracted }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const extract = async () => {
    if (!raw.trim()) return;
    setLoading(true); setError("");
    try {
      const res = await fetch("/api/extract-jd", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ raw }),
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data?.error?.message
          || (typeof data?.error === "string" ? data.error : null)
          || `API 오류 (HTTP ${res.status}) — 응답 원문: ${JSON.stringify(data).slice(0, 300)}`);
      }
      const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("\n");
      const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
      const newReqs = (parsed.requirements || []).map(r => {
        // 역량 키워드 겹침으로 매칭 경험 자동 제안 — 최종 확인은 사람이
        const match = experiences.find(e => (e.competencies || []).some(c => r.requirement.includes(c) || c.includes(r.requirement)));
        return {
          id: "r_" + Date.now() + Math.random().toString(36).slice(2, 4),
          requirement: r.requirement, category: "required_competency",
          importance: r.importance || 3, matchedExp: match ? match.id : null,
          matchReason: match ? "역량 키워드로 자동 매칭됨 — 적절한지 확인 필요" : "",
          gap: match ? "" : "매칭 가능한 경험 없음",
        };
      });
      onExtracted(newReqs);
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ marginTop: 10 }}>
      {error && (
        <div style={{ marginBottom: 8, padding: "var(--sp-3) var(--sp-4)", background: C.accent, border: `1px solid ${C.line}`, borderRadius: "var(--r-md)" }}>
          <div style={{ fontSize: "var(--fs-sm)", color: C.redText, fontWeight: 700 }}>오류</div>
          <div style={{ fontSize: "var(--fs-sm)", color: C.redText, whiteSpace: "pre-wrap", fontFamily: "monospace" }}>{error}</div>
        </div>
      )}
      <Btn small primary disabled={loading || !raw.trim()} onClick={extract}>{loading ? "추출 중…" : "AI로 요구 역량 추출"}</Btn>
    </div>
  );
}

function ApplicationReview({ app, experiences, metrics }) {
  const keywordSet = (app.requirements || []).map(r => r.requirement);
  return (
    <div>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.6 }}>
        면접·자소서 직전에 이 회사 관련 내용만 압축해서 훑어보는 화면입니다.
      </div>

      <Card style={{ marginBottom: 16 }}>
        <Label>이 회사 핵심 키워드 (공고 분석 결과)</Label>
        {keywordSet.length > 0 ? (
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {keywordSet.map(k => <Badge key={k} label={k} color={C.blueText} bg={C.blueBg} />)}
          </div>
        ) : <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>등록된 요구 역량이 없습니다. 공고 분석 탭에서 추가하세요.</div>}
      </Card>

      <Label>면접 질문별 요약</Label>
      {(app.interviews || []).map(iq => {
        const exp = experiences.find(e => e.id === iq.selectedExperienceId);
        const unverified = exp ? metrics.filter(m => m.experienceId === exp.id && m.certainty !== "verified") : [];
        const matchedKeywords = exp ? keywordSet.filter(k => (exp.competencies || []).some(c => k.includes(c) || c.includes(k))) : [];
        return (
          <Card key={iq.id} style={{ marginBottom: 10 }}>
            <div style={{ fontSize: "var(--fs-md)", fontWeight: 700, marginBottom: 8 }}>{iq.question || "(질문 미입력)"}</div>
            {exp ? (
              <>
                <div style={{ fontSize: "var(--fs-base)", color: C.text, marginBottom: 6 }}>
                  <b>{exp.title}</b> — {exp.coreMessage || exp.oneLineSummary || "핵심 메시지 미입력"}
                </div>
                {matchedKeywords.length > 0 && (
                  <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 6 }}>
                    {matchedKeywords.map(k => <Badge key={k} label={k} color={C.greenText} bg={C.greenBg} />)}
                  </div>
                )}
                {unverified.length > 0 && (
                  <div style={{ fontSize: "var(--fs-sm)", color: C.redText, background: C.redBg, padding: "var(--sp-2) var(--sp-4)", borderRadius: "var(--r-md)", marginBottom: 6 }}>
                    <CIcon icon={cilWarning} width={14} height={14} aria-hidden="true" style={{ marginRight: 5, verticalAlign: "-2px" }} />확인 안 된 수치: {unverified.map(m => `${m.metricName} (${formatMetric(m, "exact")})`).join(", ")}
                  </div>
                )}
                {(iq.followUps || []).length > 0 && (
                  <div style={{ fontSize: "var(--fs-sm)", color: C.sub }}>
                    예상 꼬리질문: {iq.followUps.join(" / ")}
                  </div>
                )}
              </>
            ) : (
              <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>사용할 경험이 아직 선택되지 않았습니다.</div>
            )}
          </Card>
        );
      })}
      {(app.interviews || []).length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.faintText }}>등록된 면접 질문이 없습니다.</div>}
    </div>
  );
}

function AppLinks({ app, setApplications }) {
  const links = app.links || [];
  const [adding, setAdding] = useState(false);
  const [label, setLabel] = useState("");
  const [url, setUrl] = useState("");

  const normalizeUrl = (u) => /^https?:\/\//i.test(u) ? u : `https://${u}`;

  const add = () => {
    if (!url.trim()) return;
    setApplications(prev => prev.map(a => a.id === app.id
      ? { ...a, links: [...(a.links || []), { id: "lk_" + Date.now(), label: label.trim() || "링크", url: normalizeUrl(url.trim()) }] } : a));
    setLabel(""); setUrl(""); setAdding(false);
  };
  const remove = (id) => setApplications(prev => prev.map(a => a.id === app.id ? { ...a, links: (a.links || []).filter(l => l.id !== id) } : a));

  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 16 }}>
      {links.map(l => (
        <span key={l.id} style={{ display: "inline-flex", alignItems: "center", gap: 5, background: C.lineSoft, borderRadius: "var(--r-lg)", padding: "var(--sp-2) var(--sp-4)", fontSize: "var(--fs-sm)" }}>
          <a href={l.url} target="_blank" rel="noreferrer" style={{ color: C.text, textDecoration: "none" }}><CIcon icon={cilLink} width={14} height={14} aria-hidden="true" style={{ marginRight: 5, verticalAlign: "-2px" }} />{l.label}</a>
          <span {...clickableProps(() => remove(l.id), { label: "닫기" })} style={{ cursor: "pointer", color: C.faintText }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
        </span>
      ))}
      {adding ? (
        <span style={{ display: "inline-flex", gap: 6, alignItems: "center" }}>
          <Input placeholder="이름 (예: 회사 홈페이지)" value={label} onChange={e => setLabel(e.target.value)} style={{ width: 140, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)" }} />
          <Input placeholder="URL" value={url} onChange={e => setUrl(e.target.value)}
            onKeyDown={e => e.key === "Enter" && add()} style={{ width: 200, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)" }} />
          <Btn small primary onClick={add}>추가</Btn>
          <Btn small onClick={() => { setAdding(false); setLabel(""); setUrl(""); }}>취소</Btn>
        </span>
      ) : (
        <span {...clickableProps(() => setAdding(true))} style={{ fontSize: "var(--fs-sm)", color: C.sub, cursor: "pointer", textDecoration: "underline" }}>+ 링크 추가 (회사 홈페이지, 채용공고, 지원 포탈 등)</span>
      )}
    </div>
  );
}

function ApplicationDetail({ app, setApplications, experiences, outputs, metrics, onBack, onOpenExp, addTrash, interviewCategories, addInterviewCategory }) {
  const [tab, setTab] = useState("공고 분석");
  const [chatEssayId, setChatEssayId] = useState(null);
  const patch = (k, v) => setApplications(prev => prev.map(a => a.id === app.id ? { ...a, [k]: v } : a));
  const deleteApp = () => {
    setApplications(prev => prev.filter(a => a.id !== app.id));
    addTrash("application", `${app.company} ${app.position}`.trim(), app);
    onBack();
  };

  return (
    <div style={{ maxWidth: 880 }}>
      <div {...clickableProps(onBack)} style={{ fontSize: "var(--fs-base)", color: C.sub, cursor: "pointer", marginBottom: 10 }}>← 지원 관리</div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 4 }} className="wrap-sm">
        <div style={{ display: "flex", gap: 8, alignItems: "baseline", flexWrap: "wrap" }} className="wrap-sm">
          <Input value={app.company} onChange={e => patch("company", e.target.value)} style={{ fontSize: "var(--fs-2xl)", fontWeight: 800, border: "none", padding: "var(--sp-1) 0", width: 220 }} />
          <Input value={app.position} placeholder="직무" onChange={e => patch("position", e.target.value)} style={{ fontSize: "var(--fs-md)", color: C.sub, border: "none", padding: "var(--sp-1) 0", width: 160 }} />
        </div>
        <span {...clickableProps(deleteApp, { label: "닫기" })} title="이 지원 삭제 (휴지통에서 복구 가능)" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-md)", padding: "var(--sp-2)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
      </div>
      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 12, fontSize: "var(--fs-base)", color: C.sub }} className="wrap-sm">
        마감 <Input value={app.deadline || ""} placeholder="YYYY-MM-DD" onChange={e => patch("deadline", e.target.value)} style={{ width: 120, border: "none", padding: "var(--sp-1) 0", color: C.sub }} />
        <span>·</span>
        우선순위
        <select value={app.priority} onChange={e => patch("priority", e.target.value)}
          style={{ fontFamily: font, fontSize: "var(--fs-base)", border: "none", background: "transparent", color: C.sub, cursor: "pointer" }}>
          <option value="high">높음</option>
          <option value="medium">보통</option>
          <option value="low">낮음</option>
        </select>
      </div>

      <AppLinks app={app} setApplications={setApplications} />

      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.line}`, marginBottom: 18 }}>
        {["공고 분석", "자소서", "면접", "복습"].map(t => (
          <div key={t} {...clickableProps(() => setTab(t))} style={{ padding: "var(--sp-3) var(--sp-5)", fontSize: "var(--fs-base)", fontWeight: tab === t ? 700 : 500, cursor: "pointer",
            color: tab === t ? C.text : C.sub, borderBottom: tab === t ? `2px solid ${C.text}` : "2px solid transparent", marginBottom: -1 }}>{t}</div>
        ))}
      </div>

      {tab === "복습" && <ApplicationReview app={app} experiences={experiences} metrics={metrics} />}

      {tab === "공고 분석" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1.3fr", gap: 16, alignItems: "start" }} className="stack-sm">
          <Card style={{ position: "sticky", top: 16 }}>
            <Label>채용공고 원문</Label>
            <Textarea rows={20} placeholder="채용공고 원문을 여기에 붙여넣으세요" value={app.jobPostingRaw || ""}
              onChange={e => setApplications(prev => prev.map(a => a.id === app.id ? { ...a, jobPostingRaw: e.target.value } : a))}
              style={{ fontSize: "var(--fs-base)", lineHeight: 1.6 }} />
            <JobPostingExtractor experiences={experiences} raw={app.jobPostingRaw || ""}
              onExtracted={(newReqs) => setApplications(prev => prev.map(a => a.id === app.id
                ? { ...a, requirements: [...a.requirements, ...newReqs] } : a))} />
          </Card>

          <Card>
            <Label>요구 역량 ↔ 경험 매칭 (추천 이유와 부족한 점 필수)</Label>
            {app.requirements.map(r => {
              const exp = experiences.find(e => e.id === r.matchedExp);
              const patchReq = (k, v) => setApplications(prev => prev.map(a => a.id === app.id
                ? { ...a, requirements: a.requirements.map(x => x.id === r.id ? { ...x, [k]: v } : x) } : a));
              const removeReq = () => {
                setApplications(prev => prev.map(a => a.id === app.id
                  ? { ...a, requirements: a.requirements.filter(x => x.id !== r.id) } : a));
                addTrash("requirement", r.requirement || "요구 역량 항목", { appId: app.id, item: r });
              };
              return (
                <div key={r.id} style={{ padding: "var(--sp-4) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                    <Input value={r.requirement} onChange={e => patchReq("requirement", e.target.value)} style={{ fontWeight: 600, border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
                    <span {...clickableProps(removeReq, { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)", flexShrink: 0, marginTop: 4 }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                  </div>
                  <div style={{ fontSize: "var(--fs-xs)", color: C.faintText, marginBottom: 6 }}>중요도 {"●".repeat(r.importance)}{"○".repeat(5 - r.importance)}</div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }} className="wrap-sm">
                    <select value={r.matchedExp || ""} onChange={e => patchReq("matchedExp", e.target.value || null)}
                      style={{ fontFamily: font, fontSize: "var(--fs-sm)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, background: C.panel, color: exp ? C.blueText : C.text }}>
                      <option value="">매칭 없음</option>
                      {experiences.map(e2 => <option key={e2.id} value={e2.id}>{e2.title}</option>)}
                    </select>
                    <Input value={r.matchReason || ""} placeholder="매칭 이유 / 부족한 점" onChange={e => patchReq("matchReason", e.target.value)}
                      style={{ fontSize: "var(--fs-sm)", color: C.sub, flex: 1, minWidth: 140 }} />
                  </div>
                </div>
              );
            })}
            {app.requirements.length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.faintText, marginBottom: 10 }}>왼쪽에 채용공고를 붙여넣고 "AI로 요구 역량 추출"을 누르거나, 아래에서 직접 추가하세요.</div>}
            <Btn small onClick={() => setApplications(prev => prev.map(a => a.id === app.id
              ? { ...a, requirements: [...a.requirements, { id: "r_" + Date.now(), requirement: "", category: "required_competency", importance: 3, matchedExp: null, matchReason: "", gap: "" }] } : a))}>
              + 요구 역량 추가
            </Btn>
          </Card>
        </div>
      )}

      {tab === "자소서" && chatEssayId && (() => {
        const q = app.essays.find(x => x.id === chatEssayId);
        return (
          <EssayChat
            title="자소서 작성 도우미"
            subtitle={q.question || "문항 미입력"}
            systemPrompt={ESSAY_COACH_SYSTEM_PROMPT}
            contextText={buildEssayContext(app, q, experiences, metrics)}
            autoStartMessage="안녕하세요, 이 문항에 사용할 경험을 추천해 주시고 초안을 작성해 주세요."
            onClose={() => setChatEssayId(null)}
            history={q.chatHistory || []}
            onHistoryChange={(h) => setApplications(prev => prev.map(a => a.id === app.id
              ? { ...a, essays: a.essays.map(x => x.id === chatEssayId ? { ...x, chatHistory: h } : x) } : a))}
            onSaveDraft={(text) => setApplications(prev => prev.map(a => a.id === app.id
              ? { ...a, essays: a.essays.map(x => x.id === chatEssayId ? { ...x, draft: text, status: "drafting" } : x) } : a))}
          />
        );
      })()}

      {tab === "자소서" && !chatEssayId && (
        <div style={{ display: "grid", gap: 12 }}>
          {app.essays.map(q => {
            const patchQ = (k, v) => setApplications(prev => prev.map(a => a.id === app.id
              ? { ...a, essays: a.essays.map(x => x.id === q.id ? { ...x, [k]: v } : x) } : a));
            const removeQ = () => {
              setApplications(prev => prev.map(a => a.id === app.id
                ? { ...a, essays: a.essays.filter(x => x.id !== q.id) } : a));
              addTrash("essay", q.question || "자소서 문항", { appId: app.id, item: q });
            };
            return (
              <Card key={q.id}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8, gap: 10 }}>
                  <Input value={q.question} onChange={e => patchQ("question", e.target.value)} style={{ fontWeight: 700, fontSize: "var(--fs-md)", border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
                  <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                    {q.isLocked && <Badge label="제출본 잠금" color={C.sub} bg={C.lineSoft} />}
                    <Badge label={{ not_started: "미시작", drafting: "초안 작성", reviewing: "검토", complete: "완료" }[q.status]}
                      color={q.status === "complete" ? C.greenText : C.blueText} bg={q.status === "complete" ? C.greenBg : C.blueBg} />
                    <span {...clickableProps(removeQ, { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 10 }}>
                  <span style={{ fontSize: "var(--fs-sm)", color: C.sub }}>글자 수 제한</span>
                  <Input type="number" value={q.characterLimit} onChange={e => patchQ("characterLimit", Number(e.target.value))} style={{ width: 80, border: "none", padding: "var(--sp-1) 0", fontSize: "var(--fs-sm)", color: C.sub }} />
                  <span style={{ fontSize: "var(--fs-sm)", color: C.sub }}>자</span>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
                  <span style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>선택 경험:</span>
                  {q.selectedExperienceIds.map(id => {
                    const e = experiences.find(x => x.id === id);
                    if (!e) return null;
                    return (
                      <span key={id} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                        <Badge label={e.title} color={C.blueText} bg={C.blueBg} />
                        <span {...clickableProps(() => patchQ("selectedExperienceIds", q.selectedExperienceIds.filter(x => x !== id)))}
                          style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-xs)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                      </span>
                    );
                  })}
                  {q.selectedExperienceIds.length === 0 && <span style={{ fontSize: "var(--fs-xs)", color: C.faintText }}>없음 — AI가 전체 경험 중 추천합니다</span>}
                  <select value="" onChange={e => {
                    const id = e.target.value;
                    if (id && !q.selectedExperienceIds.includes(id)) patchQ("selectedExperienceIds", [...q.selectedExperienceIds, id]);
                  }} style={{ fontFamily: font, fontSize: "var(--fs-xs)", padding: "var(--sp-1) var(--sp-2)", borderRadius: "var(--r-sm)", border: `1px solid ${C.line}`, background: C.panel, color: C.sub }}>
                    <option value="">+ 경험 추가</option>
                    {experiences.filter(e => !q.selectedExperienceIds.includes(e.id)).map(e => <option key={e.id} value={e.id}>{e.title}</option>)}
                  </select>
                </div>
                <div style={{ marginTop: 10 }}>
                  <Textarea value={q.draft || ""} placeholder="여기에 직접 작성해도 되고, 아래 'AI와 함께 작성'으로 도움받아도 됩니다."
                    onChange={e => patchQ("draft", e.target.value)}
                    onBlur={() => { if (q.draft?.trim() && q.status === "not_started") patchQ("status", "drafting"); }}
                    disabled={q.isLocked} rows={5} style={{ fontSize: "var(--fs-base)", lineHeight: 1.6 }} />
                  <div style={{ display: "flex", justifyContent: "flex-end", fontSize: "var(--fs-xs)", color: (q.draft || "").length > q.characterLimit ? C.redText : C.faintText, marginTop: 4 }}>
                    {(q.draft || "").length} / {q.characterLimit}자
                  </div>
                </div>
                {!q.isLocked && <div style={{ marginTop: 6 }}><Btn small onClick={() => setChatEssayId(q.id)}>AI와 함께 작성하기 →</Btn></div>}
              </Card>
            );
          })}
          <Btn small onClick={() => setApplications(prev => prev.map(a => a.id === app.id
            ? { ...a, essays: [...a.essays, { id: "q_" + Date.now(), question: "", characterLimit: 1000, status: "not_started", selectedExperienceIds: [], isLocked: false, chatHistory: [] }] } : a))}>
            + 문항 추가
          </Btn>
        </div>
      )}

      {tab === "면접" && (
        <div style={{ display: "grid", gap: 12 }}>
          {app.interviews.map(iq => {
            const exp = experiences.find(e => e.id === iq.selectedExperienceId);
            const patchIq = (k, v) => setApplications(prev => prev.map(a => a.id === app.id
              ? { ...a, interviews: a.interviews.map(x => x.id === iq.id ? { ...x, [k]: v } : x) } : a));
            const removeIq = () => {
              setApplications(prev => prev.map(a => a.id === app.id
                ? { ...a, interviews: a.interviews.filter(x => x.id !== iq.id) } : a));
              addTrash("interview", iq.question || "면접 질문", { appId: app.id, item: iq });
            };
            return (
              <Card key={iq.id}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, gap: 10 }} className="wrap-sm">
                  <Input value={iq.question} onChange={e => patchIq("question", e.target.value)} style={{ fontWeight: 700, fontSize: "var(--fs-md)", border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
                  <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                    <CategorySelect value={iq.category} options={interviewCategories} placeholder="카테고리 없음" onAddOption={addInterviewCategory} onChange={(v) => patchIq("category", v)} />
                    <span {...clickableProps(removeIq, { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: "var(--fs-base)", color: C.sub }}>사용 경험:</span>
                  <select value={iq.selectedExperienceId || ""} onChange={e => patchIq("selectedExperienceId", e.target.value || null)}
                    style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-lg)", border: `1px solid ${C.line}`, background: C.panel, color: exp ? C.blueText : C.text }}>
                    <option value="">선택 안 함</option>
                    {experiences.map(e2 => <option key={e2.id} value={e2.id}>{e2.title}</option>)}
                  </select>
                  {exp && <span style={{ fontSize: "var(--fs-sm)", color: C.sub }}>연습 {iq.practiceCount}회 · 자신감 {iq.confidence ?? "—"}/5</span>}
                </div>
                {exp ? (
                  <>
                    <Label>예상 꼬리질문</Label>
                    {iq.followUps.map((f, i) => (
                      <div key={i} style={{ display: "flex", gap: 6, alignItems: "center", padding: "var(--sp-1) 0" }}>
                        <span style={{ fontSize: "var(--fs-base)", color: C.sub, flex: 1 }}>· {f}</span>
                        <span {...clickableProps(() => patchIq("followUps", iq.followUps.filter((_, fi) => fi !== i)))} style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-xs)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                      </div>
                    ))}
                    <div style={{ marginTop: 10, display: "flex", gap: 8 }} className="wrap-sm">
                      <Btn small primary disabled title="준비 중인 기능입니다">60초 연습 시작</Btn><Btn small disabled title="준비 중인 기능입니다">키워드 가리기</Btn>
                    </div>
                  </>
                ) : (
                  <div style={{ fontSize: "var(--fs-base)", color: C.orangeText, background: C.accent, border: `1px solid ${C.line}`, padding: "var(--sp-4) var(--sp-5)", borderRadius: "var(--r-lg)" }}>
                    사용할 경험이 선택되지 않았습니다.
                  </div>
                )}
              </Card>
            );
          })}
          <Btn small onClick={() => setApplications(prev => prev.map(a => a.id === app.id
            ? { ...a, interviews: [...a.interviews, { id: "iq_" + Date.now(), question: "", category: "achievement", selectedExperienceId: null, practiceCount: 0, confidence: null, followUps: [] }] } : a))}>
            + 질문 추가
          </Btn>
        </div>
      )}
    </div>
  );
}

/* ============================================================ 기본 이력서 */
/* ============================================================ 휴지통 */
function formatDeletedAt(iso) {
  try {
    const d = new Date(iso);
    const diffMin = Math.round((Date.now() - d.getTime()) / 60000);
    if (diffMin < 1) return "방금";
    if (diffMin < 60) return `${diffMin}분 전`;
    const diffHr = Math.round(diffMin / 60);
    if (diffHr < 24) return `${diffHr}시간 전`;
    return d.toISOString().slice(0, 16).replace("T", " ");
  } catch { return ""; }
}

function Trash({ trash, onRestore, onPurge, onClear }) {
  const typeLabel = { experience: "경험", skill: "스킬", cert: "자격증", award: "수상기록", application: "지원", requirement: "요구 역량", essay: "자소서 문항", interview: "면접 질문", timeline_activity: "타임라인 활동" };
  return (
    <div style={{ maxWidth: 700 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <H2>휴지통</H2>
        {trash.length > 0 && <Btn small onClick={onClear}>전체 비우기</Btn>}
      </div>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.6 }}>
        삭제한 항목은 영구 삭제하기 전까지 여기서 복구할 수 있습니다.
      </div>
      {trash.length === 0 && (
        <Card><div style={{ fontSize: "var(--fs-base)", color: C.faintText, textAlign: "center", padding: "var(--sp-5) 0" }}>삭제한 항목이 없습니다.</div></Card>
      )}
      {trash.map(t => (
        <Card key={t.id} style={{ marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <Badge label={typeLabel[t.type] || t.type} color={C.sub} bg={C.lineSoft} />
            <span style={{ marginLeft: 8, fontSize: "var(--fs-base)", fontWeight: 600 }}>{t.label || "(제목 없음)"}</span>
            <div style={{ fontSize: "var(--fs-xs)", color: C.faintText, marginTop: 3 }}>{formatDeletedAt(t.deletedAt)} 삭제됨</div>
          </div>
          <div style={{ display: "flex", gap: 8, flexShrink: 0 }} className="wrap-sm">
            <Btn small primary onClick={() => onRestore(t.id)}>복구</Btn>
            <Btn small onClick={() => onPurge(t.id)}>영구 삭제</Btn>
          </div>
        </Card>
      ))}
    </div>
  );
}

/* ============================================================ 마스터 자소서·면접 */
function MasterPrep({ essays, setEssays, interviews, setInterviews, experiences, metrics, resumeProfile, interviewCategories, addInterviewCategory }) {
  const [tab, setTab] = useState("자소서");
  const [chatId, setChatId] = useState(null);
  const [iqChatId, setIqChatId] = useState(null);
  const masterApp = { company: "마스터 (공통 준비용)", position: resumeProfile.targetRole || "", requirements: [] };

  const patchQ = (id, k, v) => setEssays(prev => prev.map(x => x.id === id ? { ...x, [k]: v } : x));
  const removeQ = (id) => setEssays(prev => prev.filter(x => x.id !== id));
  const addQ = () => setEssays(prev => [...prev, { id: "mq_" + Date.now(), question: "", characterLimit: 1000, draft: "", status: "not_started", chatHistory: [] }]);

  const patchIq = (id, k, v) => setInterviews(prev => prev.map(x => x.id === id ? { ...x, [k]: v } : x));
  const removeIq = (id) => setInterviews(prev => prev.filter(x => x.id !== id));
  const addIq = () => setInterviews(prev => [...prev, { id: "miq_" + Date.now(), question: "", category: "", selectedExperienceId: null, practiceCount: 0, confidence: null, followUps: [], draft: "", chatHistory: [] }]);

  return (
    <div style={{ maxWidth: 820 }}>
      <H2>자소서·면접 준비 (공통)</H2>
      <div style={{ fontSize: "var(--fs-base)", color: C.sub, marginBottom: 16, lineHeight: 1.6 }}>
        특정 회사에 매지 않고, 자주 나오는 공통 문항을 미리 준비해두는 곳입니다. 여기서 만든 답변은 지원 관리의 각 회사별 문항을 쓸 때 참고용으로 활용하세요.
      </div>

      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.line}`, marginBottom: 18 }}>
        {["자소서", "면접"].map(t => (
          <div key={t} {...clickableProps(() => setTab(t))} style={{ padding: "var(--sp-3) var(--sp-5)", fontSize: "var(--fs-base)", fontWeight: tab === t ? 700 : 500, cursor: "pointer",
            color: tab === t ? C.text : C.sub, borderBottom: tab === t ? `2px solid ${C.text}` : "2px solid transparent", marginBottom: -1 }}>{t}</div>
        ))}
      </div>

      {tab === "자소서" && chatId && (() => {
        const q = essays.find(x => x.id === chatId);
        return (
          <EssayChat
            title="자소서 작성 도우미"
            subtitle={q.question || "문항 미입력"}
            systemPrompt={ESSAY_COACH_SYSTEM_PROMPT}
            contextText={buildEssayContext(masterApp, q, experiences, metrics)}
            autoStartMessage="안녕하세요, 이 문항에 사용할 경험을 추천해 주시고 초안을 작성해 주세요."
            onClose={() => setChatId(null)}
            history={q.chatHistory || []}
            onHistoryChange={(h) => setEssays(prev => prev.map(x => x.id === chatId ? { ...x, chatHistory: h } : x))}
            onSaveDraft={(text) => setEssays(prev => prev.map(x => x.id === chatId ? { ...x, draft: text, status: "drafting" } : x))}
          />
        );
      })()}

      {tab === "자소서" && !chatId && (
        <div style={{ display: "grid", gap: 12 }}>
          {essays.map(q => (
            <Card key={q.id}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8, gap: 10 }}>
                <Input value={q.question} onChange={e => patchQ(q.id, "question", e.target.value)} style={{ fontWeight: 700, fontSize: "var(--fs-md)", border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
                <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                  <Badge label={{ not_started: "미시작", drafting: "초안 작성", reviewing: "검토", complete: "완료" }[q.status]}
                    color={q.status === "complete" ? C.greenText : C.blueText} bg={q.status === "complete" ? C.greenBg : C.blueBg} />
                  <span {...clickableProps(() => removeQ(q.id), { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                </div>
              </div>
              <Textarea value={q.draft || ""} placeholder="여기에 직접 작성해도 되고, 아래 'AI와 함께 작성'으로 도움받아도 됩니다."
                onChange={e => patchQ(q.id, "draft", e.target.value)}
                onBlur={() => { if (q.draft?.trim() && q.status === "not_started") patchQ(q.id, "status", "drafting"); }}
                rows={5} style={{ fontSize: "var(--fs-base)", lineHeight: 1.6 }} />
              <div style={{ display: "flex", justifyContent: "flex-end", fontSize: "var(--fs-xs)", color: (q.draft || "").length > q.characterLimit ? C.redText : C.faintText, marginTop: 4, marginBottom: 8 }}>
                {(q.draft || "").length} / {q.characterLimit}자
              </div>
              <Btn small onClick={() => setChatId(q.id)}>AI와 함께 작성하기 →</Btn>
            </Card>
          ))}
          <Btn small onClick={addQ}>+ 문항 추가</Btn>
        </div>
      )}

      {tab === "면접" && iqChatId && (() => {
        const iq = interviews.find(x => x.id === iqChatId);
        const pseudoEssay = { question: iq.question, characterLimit: 400 };
        return (
          <EssayChat
            title="면접 답변 작성 도우미"
            subtitle={iq.question || "질문 미입력"}
            systemPrompt={ESSAY_COACH_SYSTEM_PROMPT}
            contextText={buildEssayContext(masterApp, pseudoEssay, experiences, metrics)}
            autoStartMessage="안녕하세요, 이 면접 질문에 쓸 경험을 추천해 주시고 답변 초안을 작성해 주세요."
            saveDraftLabel="이 답변을 초안으로 저장"
            onClose={() => setIqChatId(null)}
            history={iq.chatHistory || []}
            onHistoryChange={(h) => setInterviews(prev => prev.map(x => x.id === iqChatId ? { ...x, chatHistory: h } : x))}
            onSaveDraft={(text) => setInterviews(prev => prev.map(x => x.id === iqChatId ? { ...x, draft: text } : x))}
          />
        );
      })()}

      {tab === "면접" && !iqChatId && (
        <div style={{ display: "grid", gap: 12 }}>
          {interviews.map(iq => {
            const exp = experiences.find(e => e.id === iq.selectedExperienceId);
            return (
              <Card key={iq.id}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, gap: 10 }} className="wrap-sm">
                  <Input value={iq.question} onChange={e => patchIq(iq.id, "question", e.target.value)} style={{ fontWeight: 700, fontSize: "var(--fs-md)", border: "none", padding: "var(--sp-1) 0", flex: 1 }} />
                  <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                    <CategorySelect value={iq.category} options={interviewCategories} placeholder="카테고리 없음"
                      onAddOption={addInterviewCategory} onChange={(v) => patchIq(iq.id, "category", v)} />
                    <span {...clickableProps(() => removeIq(iq.id), { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-base)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: "var(--fs-base)", color: C.sub }}>참고 경험:</span>
                  <select value={iq.selectedExperienceId || ""} onChange={e => patchIq(iq.id, "selectedExperienceId", e.target.value || null)}
                    style={{ fontFamily: font, fontSize: "var(--fs-base)", padding: "var(--sp-2) var(--sp-3)", borderRadius: "var(--r-md)", border: `1px solid ${C.line}`, background: C.panel }}>
                    <option value="">선택 안 함</option>
                    {experiences.map(e2 => <option key={e2.id} value={e2.id}>{e2.title}</option>)}
                  </select>
                  {exp && <span style={{ fontSize: "var(--fs-sm)", color: C.sub }}>연습 {iq.practiceCount}회</span>}
                </div>
                {iq.draft && (
                  <div style={{ padding: "var(--sp-4) var(--sp-5)", background: C.bg, border: `1px solid ${C.line}`, borderRadius: "var(--r-md)", fontSize: "var(--fs-sm)", color: C.sub, lineHeight: 1.6, marginBottom: 10, maxHeight: 90, overflow: "hidden" }}>
                    {iq.draft}
                  </div>
                )}
                <Btn small onClick={() => setIqChatId(iq.id)}>{iq.draft ? "답변 이어서 작성하기 →" : "답변 작성 도우미 열기 →"}</Btn>
              </Card>
            );
          })}
          <Btn small onClick={addIq}>+ 질문 추가</Btn>
        </div>
      )}
    </div>
  );
}

function Resume({ experiences, outputs, metrics, resumeProfile, setResumeProfile, skills, certs, setCerts, awards, setAwards, addTrash }) {
  const approved = outputs.filter(o => o.outputType === "resume" && o.approvalStatus === "approved");
  const patch = (k, v) => setResumeProfile(p => ({ ...p, [k]: v }));
  const autosave = useAutosave(JSON.stringify(resumeProfile));

  const [certDraft, setCertDraft] = useState({ name: "", issuer: "", date: "", note: "" });
  const patchCert = (id, k, v) => setCerts(prev => prev.map(c => c.id === id ? { ...c, [k]: v } : c));
  const addCert = () => {
    if (!certDraft.name.trim()) return;
    setCerts(prev => [...prev, { id: "c_" + Date.now(), ...certDraft }]);
    setCertDraft({ name: "", issuer: "", date: "", note: "" });
  };
  const removeCert = (id) => {
    const cert = certs.find(c => c.id === id);
    setCerts(prev => prev.filter(c => c.id !== id));
    addTrash("cert", cert.name, cert);
  };

  const [awardDraft, setAwardDraft] = useState({ name: "", issuer: "", date: "", note: "" });
  const patchAward = (id, k, v) => setAwards(prev => prev.map(a => a.id === id ? { ...a, [k]: v } : a));
  const addAward = () => {
    if (!awardDraft.name.trim()) return;
    setAwards(prev => [...prev, { id: "aw_" + Date.now(), ...awardDraft }]);
    setAwardDraft({ name: "", issuer: "", date: "", note: "" });
  };
  const removeAward = (id) => {
    const award = awards.find(a => a.id === id);
    setAwards(prev => prev.filter(a => a.id !== id));
    addTrash("award", award.name, award);
  };

  // 승인된 문장을 경험(소속·역할) 단위로 그룹핑
  const groups = approved.reduce((acc, o) => {
    const exp = experiences.find(e => e.id === o.experienceId);
    const key = exp ? `${exp.organization} · ${exp.role}` : "기타";
    (acc[key] = acc[key] || []).push(o);
    return acc;
  }, {});

  // 근거가 연결된 스킬 항목만 배지로 노출
  const linkedSkillBadges = (skills || []).flatMap(s =>
    (s.scopeItems || []).filter(it => it.evidenceExpId).slice(0, 1).map(it => `${s.name} · ${it.text}`)
  );

  const exportWord = () => {
    const careerHtml = Object.entries(groups).map(([label, items]) => `
      <h3 style="font-size:14px;margin:14px 0 4px;">${label}</h3>
      <ul style="margin:0 0 8px 0; padding-left:18px;">
        ${items.map(o => `<li style="margin-bottom:4px;">${resolveTokenText(o.content, metrics)}</li>`).join("")}
      </ul>`).join("") || "<p>등록된 경력 문장이 없습니다.</p>";

    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="utf-8"><title>이력서</title></head>
      <body style="font-family:'맑은 고딕',sans-serif; font-size:13px; color:#222;">
        <h1 style="font-size:22px; margin-bottom:4px;">${resumeProfile.name || "이름 미입력"}</h1>
        <p style="color:#555; margin:0 0 4px;">${resumeProfile.targetRole || ""}</p>
        <p style="color:#555; margin:0 0 16px;">${[resumeProfile.email, resumeProfile.phone].filter(Boolean).join(" · ")}</p>
        <p style="margin:0 0 16px;">${resumeProfile.headline || ""}</p>
        <h2 style="font-size:16px; border-bottom:1px solid #ccc; padding-bottom:4px;">경력</h2>
        ${careerHtml}
        <h2 style="font-size:16px; border-bottom:1px solid #ccc; padding-bottom:4px; margin-top:20px;">역량</h2>
        <p>${linkedSkillBadges.join(", ") || "등록된 역량이 없습니다."}</p>
        <h2 style="font-size:16px; border-bottom:1px solid #ccc; padding-bottom:4px; margin-top:20px;">자격증 · 어학</h2>
        <ul style="padding-left:18px;">${(certs.length ? certs : [{ name: "등록된 자격증이 없습니다." }]).map(c => `<li>${c.name}${c.issuer ? ` · ${c.issuer}` : ""}${c.date ? ` · ${c.date}` : ""}</li>`).join("")}</ul>
        <h2 style="font-size:16px; border-bottom:1px solid #ccc; padding-bottom:4px; margin-top:20px;">수상기록</h2>
        <ul style="padding-left:18px;">${(awards.length ? awards : [{ name: "등록된 수상기록이 없습니다." }]).map(a => `<li>${a.name}${a.issuer ? ` · ${a.issuer}` : ""}${a.date ? ` · ${a.date}` : ""}</li>`).join("")}</ul>
      </body></html>`;

    const blob = new Blob(["\ufeff", html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `이력서_${resumeProfile.name || "career_os"}.doc`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ maxWidth: 700 }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; left: 0; top: 0; width: 100%; }
        }
      `}</style>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
        <H2>기본 이력서</H2>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <AutosaveIndicator state={autosave} />
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }} className="wrap-sm">
        <div style={{ fontSize: "var(--fs-base)", color: C.sub }}>경력 문장은 직접 입력하지 않고, 경험 보관함의 <b>승인된</b> 문장만 불러옵니다.</div>
        <div style={{ display: "flex", gap: 6, flexShrink: 0 }} className="wrap-sm">
          <Btn small onClick={() => window.print()}>PDF로 저장 (인쇄)</Btn>
          <Btn small onClick={exportWord}>Word로 내보내기</Btn>
        </div>
      </div>
      <div className="print-area">
      <Card style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <Label>기본 정보</Label>
          {resumeProfile._imported && <Badge label="가져온 항목 · 확인 필요" color={C.orangeText} bg={C.orangeBg} />}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 6 }} className="stack-sm">
          <div><Label>이름</Label><Input placeholder="이름" value={resumeProfile.name || ""} onChange={e => patch("name", e.target.value)} /></div>
          <div><Label>희망 직무</Label><Input placeholder="예: MD / 이커머스" value={resumeProfile.targetRole || ""} onChange={e => patch("targetRole", e.target.value)} /></div>
          <div><Label>이메일</Label><Input type="email" inputMode="email" autoComplete="email" placeholder="이메일" value={resumeProfile.email || ""} onChange={e => patch("email", e.target.value)} /></div>
          <div><Label>연락처</Label><Input placeholder="연락처" value={resumeProfile.phone || ""} onChange={e => patch("phone", e.target.value)} /></div>
          <div style={{ gridColumn: "1 / -1" }}><Label>한 줄 소개</Label><Input placeholder="한 줄 소개" value={resumeProfile.headline || ""} onChange={e => patch("headline", e.target.value)} /></div>
        </div>
      </Card>
      <Card>
        <Label>경력</Label>
        {Object.keys(groups).length === 0 && <div style={{ fontSize: "var(--fs-base)", color: C.faintText, padding: "var(--sp-3) 0" }}>승인된 경력 문장이 없습니다. 경험 상세의 「활용 문장」 탭에서 문장을 승인하면 여기에 표시됩니다.</div>}
        {Object.entries(groups).map(([label, items]) => (
          <div key={label} style={{ marginBottom: 10 }}>
            <div style={{ fontSize: "var(--fs-sm)", fontWeight: 700, color: C.sub, marginBottom: 4 }}>{label}</div>
            {items.map(o => (
              <div key={o.id} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "var(--sp-4) 0", borderBottom: `1px solid ${C.lineSoft}` }}>
                <span style={{ color: C.faintText }}>·</span>
                <span style={{ fontSize: "var(--fs-base)", lineHeight: 1.6, flex: 1 }}><TokenText text={o.content} metrics={metrics} /></span>
                <Badge label="승인됨" color={C.greenText} bg={C.greenBg} />
              </div>
            ))}
          </div>
        ))}
        <div style={{ marginTop: 12 }}><Btn small disabled title="준비 중인 기능입니다">+ 경험 보관함에서 문장 불러오기</Btn></div>
        <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, marginTop: 10 }}>미승인(AI 초안) 문장은 여기에 표시되지 않습니다.</div>
      </Card>
      <Card style={{ marginTop: 12 }}>
        <Label>역량</Label>
        <div style={{ fontSize: "var(--fs-base)", color: C.sub, lineHeight: 1.6 }}>
          역량·스킬 탭에서 <b>경험 근거가 연결된 항목</b>만 자동으로 불러옵니다. (근거 없는 항목은 제외)
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 10 }}>
          {linkedSkillBadges.length === 0
            ? <div style={{ fontSize: "var(--fs-sm)", color: C.faintText }}>근거가 연결된 역량이 아직 없습니다.</div>
            : linkedSkillBadges.map(t => <Badge key={t} label={t} color={C.sub} bg={C.lineSoft} />)}
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <Label>자격증 · 어학</Label>
        {certs.map(c => (
          <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 110px 1fr 20px", gap: 10, alignItems: "center", padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }} className="tbl-row">
            <Input value={c.name} onChange={e => patchCert(c.id, "name", e.target.value)} style={{ fontWeight: 600, border: "none", padding: "var(--sp-1) 0" }} />
            <Input value={c.issuer || ""} placeholder="발급 기관" onChange={e => patchCert(c.id, "issuer", e.target.value)} style={{ color: C.sub, fontSize: "var(--fs-sm)", border: "none", padding: "var(--sp-1) 0" }} />
            <Input value={c.date || ""} placeholder="취득일" onChange={e => patchCert(c.id, "date", e.target.value)} style={{ color: C.sub, fontSize: "var(--fs-sm)", border: "none", padding: "var(--sp-1) 0" }} />
            <Input value={c.note || ""} placeholder="비고" onChange={e => patchCert(c.id, "note", e.target.value)} style={{ color: C.faintText, fontSize: "var(--fs-sm)", border: "none", padding: "var(--sp-1) 0" }} />
            <span {...clickableProps(() => removeCert(c.id), { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
          </div>
        ))}
        {certs.length === 0 && <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, padding: "var(--sp-2) 0" }}>등록된 자격증·어학이 없습니다.</div>}
        <div style={{ display: "flex", gap: 8, marginTop: 12 }} className="wrap-sm">
          <Input placeholder="자격증명" value={certDraft.name} onChange={e => setCertDraft(d => ({ ...d, name: e.target.value }))} style={{ flex: 1.3 }} />
          <Input placeholder="발급 기관" value={certDraft.issuer} onChange={e => setCertDraft(d => ({ ...d, issuer: e.target.value }))} style={{ flex: 1 }} />
          <Input placeholder="취득일 (YYYY-MM)" value={certDraft.date} onChange={e => setCertDraft(d => ({ ...d, date: e.target.value }))} style={{ width: 130 }} />
          <Btn small onClick={addCert}>추가</Btn>
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <Label>수상기록</Label>
        {awards.map(a => (
          <div key={a.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 110px 1fr 20px", gap: 10, alignItems: "center", padding: "var(--sp-3) 0", borderBottom: `1px solid ${C.lineSoft}`, fontSize: "var(--fs-base)" }} className="tbl-row">
            <Input value={a.name} onChange={e => patchAward(a.id, "name", e.target.value)} style={{ fontWeight: 600, border: "none", padding: "var(--sp-1) 0" }} />
            <Input value={a.issuer || ""} placeholder="수여 기관" onChange={e => patchAward(a.id, "issuer", e.target.value)} style={{ color: C.sub, fontSize: "var(--fs-sm)", border: "none", padding: "var(--sp-1) 0" }} />
            <Input value={a.date || ""} placeholder="수상일" onChange={e => patchAward(a.id, "date", e.target.value)} style={{ color: C.sub, fontSize: "var(--fs-sm)", border: "none", padding: "var(--sp-1) 0" }} />
            <Input value={a.note || ""} placeholder="비고 (예: 대상, 참가팀 30개 중 1위)" onChange={e => patchAward(a.id, "note", e.target.value)} style={{ color: C.faintText, fontSize: "var(--fs-sm)", border: "none", padding: "var(--sp-1) 0" }} />
            <span {...clickableProps(() => removeAward(a.id), { label: "닫기" })} title="삭제" style={{ cursor: "pointer", color: C.faintText, fontSize: "var(--fs-sm)" }}><CIcon icon={cilX} width={14} height={14} aria-hidden="true" /></span>
          </div>
        ))}
        {awards.length === 0 && <div style={{ fontSize: "var(--fs-sm)", color: C.faintText, padding: "var(--sp-2) 0" }}>등록된 수상기록이 없습니다.</div>}
        <div style={{ display: "flex", gap: 8, marginTop: 12 }} className="wrap-sm">
          <Input placeholder="수상명 (예: 전국 대학생 공모전 대상)" value={awardDraft.name} onChange={e => setAwardDraft(d => ({ ...d, name: e.target.value }))} style={{ flex: 1.3 }} />
          <Input placeholder="수여 기관" value={awardDraft.issuer} onChange={e => setAwardDraft(d => ({ ...d, issuer: e.target.value }))} style={{ flex: 1 }} />
          <Input placeholder="수상일 (YYYY-MM)" value={awardDraft.date} onChange={e => setAwardDraft(d => ({ ...d, date: e.target.value }))} style={{ width: 130 }} />
          <Btn small onClick={addAward}>추가</Btn>
        </div>
      </Card>
      </div>
    </div>
  );
}
