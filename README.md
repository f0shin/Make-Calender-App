# Make_Calender_Automation_App

## 目次

- [プロジェクト概要](#プロジェクト概要)
- [技術スタック](#技術スタック)
- [開発環境](#開発環境)
- [主要機能](#主要機能)
- [設計概要](#設計概要)
- [学んだこと](#学んだこと)
- [改善点](#改善点)

<br>

## プロジェクト概要

VBAとMS Officeを活用したキャレンダー制作プログラムの開発<br>
<br>

## 技術スタック

#VBA, #MS_Office_Excel, #HTMLDocument<br>
<br>

## 開発環境

✅ 言語 :VBA<br>
✅ OS : Windows<br>
✅ Editor : MS Office Excel<br>
✅ インタープリタ方式<br>
<br>

## 主要機能

✅ エクセルの自動化とUI管理<br>
✅ HTTPリクエストによりウェブページから祝日データを取得<br>
✅ エクセルのセキュリティと管理者モード<br>
<br>

## 設計概要

1️⃣ AcativeXオブジェクトからの入力と設定<br>
✔️ ユーザーからの入力(shtMain.txtYear, shtMain.txtMonth, shtMain.opCountryKR)で、requestYear, requestMonth, requestCountryの値を設定。<br>
<br>
2️⃣ 祝日データの持ち込み　(Get_Holiday())<br>
✔️ MSXML2.XMLHTTPを使用して timeanddate.com で祝日データクローリング。<br>
✔️ HTMLDocumentを活用してholidays-tableからデータを抽出。<br>
✔️ arrHoliday配列に保存。<br>
<br>
3️⃣ カレンダー作成　(Make_Calender())<br>
✔️ エクスポートシート(基本フォーム)をコピーして新しいファイルを作成。<br>
✔️ 日付別セル配置。(current Row、current Col)<br>
✔️ 土曜日（青）、日曜日・祝日（赤）適用。<br>
✔️ 6週目が必要な場合は動的行を追加。<br>
<br>
4️⃣ 管理者・ユーザーモードの実現<br>
✔️ btn Admin_Click()によってモードの切り替え<br>
✔️ セルに関する作業・オブジェクトの移動を制限。　(LockCells And Shapes())<br>
✔️ 管理者モードのみ保護解除。　(UnlockCellsAndShapes())<br>
<br>

## 学んだこと

✅ 最適化されたロジック及び安定性の向上のための完結したコーディングについての知識<br>
✅ ウェブクローリングについての知識<br>
<br>

## 改善点

😥 祝日データのローディングの最適化 ・・・ Get_Holiday()でクローリングに失敗すると全体の流れに影響が生じて、以後のロジックが正常に作動しない問題が発生する。 <br>
😊 クロリングの失敗時、arrHolidayを空配列に維持し自然に進行させる。<br>
<br>
😥 管理者モードのUX改善 ・・・ btnAdminを押すとき、管理者モードの活性化の可否を直観的に表示しないため、誤って移動してしまう場合が発生する。 <br>
😊 ボタンのキャプションの値を変更して現在のモードを表示させる。<br>

<br>
<br>

## 목차

- [프로젝트 개요](#프로젝트-개요)
- [기술 스택](#기술-스택)
- [개발 환경](#개발-환경)
- [주요 기능](#주요-기능)
- [설계 개요](#설계-개요)
- [배운 점](#배운-점)
- [아쉬운 점 및 개선 방안](#아쉬운-점-및-개선-방안)

<br>

## 프로젝트 개요

VBA와 MS Office를 활용한 달력 만들기 프로그램 개발<br>
<br>

## 기술 스택

#VBA, #MS_Office_Excel, #HTMLDocument<br>
<br>

## 개발 환경

✅ 언어: VBA<br>
✅ OS: Windows<br>
✅ Editor: MS Office Excel<br>
✅ 인터프리터 방식<br>
<br>

## 주요 기능

✅ 엑셀 자동화 및 UI 관리<br>
✅ HTTP 요청을 통해 웹 페이지에서 공휴일 데이터 가져오기<br>
✅ 엑셀 보안 및 관리자 모드<br>
<br>

## 설계 개요

1️⃣ AcativeX 객체를 통한 입력 및 설정<br>
✔️ 사용자 입력 (shtMain.txtYear, shtMain.txtMonth, shtMain.opCountryKR) <br>
✔️ requestYear, requestMonth, requestCountry 값 설정<br>
<br>
2️⃣ 공휴일 데이터 가져오기(Get_Holiday())<br>
✔️ MSXML2.XMLHTTP를 사용하여 timeanddate.com에서 공휴일 데이터 크롤링 <br>
✔️ HTMLDocument를 활용해 holidays-table에서 데이터를 추출 <br>
✔️ arrHoliday 배열에 저장<br>
<br>
3️⃣ 달력 생성(Make_Calender())<br>
✔️ export 시트(기본 양식)를 복사하여 새 파일 생성 <br>
✔️ 날짜별 셀 배치 (currentRow, currentCol) <br>
✔️ 토요일(파란색), 일요일·공휴일(빨간색) 적용 <br>
✔️ 6주차가 필요한 경우 동적 행 추가<br>
<br>
4️⃣ 관리자 & 사용자 모드 구현<br>
✔️ btnAdmin_Click()을 통해 모드 토글 <br>
✔️ 셀 잠금 및 도형 이동 제한 (LockCellsAndShapes()) <br>
✔️ 관리자 모드일 때만 보호 해제 (UnlockCellsAndShapes())<br>
<br>

## 배운 점

✅ 최적화된 로직 & 안정성 향상을 위해 깔끔한 코딩에 대해 많이 찾아봄.<br>
✅ 웹 크롤링에 대한 지식.<br>
<br>

## 아쉬운 점 및 개선 방안

😥 공휴일 데이터 로딩 최적화 - Get_Holiday()에서 크롤링 실패하면 전체 흐름에 영향이 생겨 이후 로직이 정상 작동하지 않는 문제가 발생함. <br>
😊 크롤링 실패 시 arrHoliday를 빈 배열로 유지하여 자연스럽게 진행함.<br>
<br>
😥 관리자 모드 UX 개선 - btnAdmin을 누를 때 관리자 모드 활성화 여부를 직관적으로 표시하지 않아서 실수로 이동하게 되는 경우가 발생함. <br>
😊 버튼의 캡션을 변경하여 현재 모드 표시함.<br>
<br>
