const CONFIG = {
  courseTitle: '주력 강의 재설계 과정',
  courseSubtitle: '자신의 강의를 한 단계 업그레이드하고 싶은 강사를 위한 2일 공개과정',
  adminEmail: 'JongmokJ@gmail.com',
  formTitle: '주력 강의 재설계 과정 신청서',
  responseSheetTitle: '주력 강의 재설계 과정_신청응답',
  confirmationMessage:
    '신청이 접수되었습니다. 신청 내용 확인 후 등록 및 결제 안내를 순차적으로 드리겠습니다.',
  introLines: [
    '주력 강의 재설계 과정 신청서입니다.',
    '참가 신청 후 신청 내용 확인을 거쳐 등록 및 결제 안내를 드립니다.',
    '참가비는 110만원(부가세 포함)이며, 교재·식사·음료 및 다과가 제공됩니다.'
  ]
};

function createCourseForm() {
  const form = FormApp.create(CONFIG.formTitle);
  form.setDescription(CONFIG.introLines.join('\n'));
  form.setCollectEmail(true);
  form.setAllowResponseEdits(false);
  form.setConfirmationMessage(CONFIG.confirmationMessage);
  form.setPublished(true);

  addParticipantSection_(form);
  addCourseSection_(form);
  addImprovementSection_(form);
  addMaterialSection_(form);
  addAgreementSection_(form);

  const sheet = SpreadsheetApp.create(CONFIG.responseSheetTitle);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  PropertiesService.getScriptProperties().setProperties({
    FORM_ID: form.getId(),
    FORM_URL: form.getPublishedUrl(),
    FORM_EDIT_URL: form.getEditUrl(),
    SHEET_ID: sheet.getId(),
    SHEET_URL: sheet.getUrl()
  });

  installFormSubmitTrigger_(form);
  formatResponseSheet_(sheet);
  logProjectLinks_(form, sheet);
  notifyProjectCreated_(form, sheet);

  return {
    formUrl: form.getPublishedUrl(),
    formEditUrl: form.getEditUrl(),
    sheetUrl: sheet.getUrl()
  };
}

function addParticipantSection_(form) {
  form.addSectionHeaderItem()
    .setTitle('1. 신청자 기본 정보')
    .setHelpText('등록 및 안내를 위한 기본 정보입니다.');

  form.addTextItem()
    .setTitle('이름')
    .setRequired(true);

  form.addTextItem()
    .setTitle('소속')
    .setRequired(true);

  form.addTextItem()
    .setTitle('직무 또는 역할')
    .setRequired(true);

  form.addTextItem()
    .setTitle('연락처')
    .setRequired(true);
}

function addCourseSection_(form) {
  form.addSectionHeaderItem()
    .setTitle('2. 주력 강의 정보')
    .setHelpText('현재 운영 중인 강의를 기준으로 확인합니다.');

  form.addTextItem()
    .setTitle('주력 강의명')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('강의 주제 또는 핵심 내용')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('주요 학습자 또는 청중')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('현재 강의 운영 상태')
    .setChoiceValues([
      '현재 정기적으로 운영 중',
      '간헐적으로 운영 중',
      '운영 경험은 있으나 최근에는 진행하지 않음'
    ])
    .setRequired(true);
}

function addImprovementSection_(form) {
  form.addSectionHeaderItem()
    .setTitle('3. 우선 개선 과제')
    .setHelpText('현재 가장 먼저 정리하거나 수정하고 싶은 항목을 선택해 주세요.');

  const item = form.addCheckboxItem()
    .setTitle('우선 개선 항목')
    .setChoiceValues([
      '핵심 메시지 정리',
      '강의 구조 재구성',
      '학습 목표 정리',
      '질문 설계',
      '활동 구성',
      '학습자 참여 유도',
      '전달기법 점검',
      '강의 자료 구성'
    ])
    .setRequired(true);

  item.setValidation(
    FormApp.createCheckboxValidation()
      .requireSelectAtLeast(1)
      .build()
  );

  form.addParagraphTextItem()
    .setTitle('현재 가장 먼저 해결하고 싶은 문제')
    .setRequired(true);
}

function addMaterialSection_(form) {
  form.addSectionHeaderItem()
    .setTitle('4. 제출 가능 자료')
    .setHelpText('사전 검토 시 활용 가능한 자료를 확인합니다.');

  form.addCheckboxItem()
    .setTitle('제출 가능한 자료')
    .setChoiceValues([
      '강의 개요',
      '강의안',
      '슬라이드',
      '워크북 또는 교재',
      '기타 참고 자료'
    ])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('사전 검토 시 참고할 사항')
    .setRequired(false);

  form.addTextItem()
    .setTitle('추천인 성함')
    .setRequired(false);
}

function addAgreementSection_(form) {
  form.addSectionHeaderItem()
    .setTitle('5. 확인 사항')
    .setHelpText('아래 내용을 확인한 뒤 제출해 주세요.');

  const agreement = form.addCheckboxItem()
    .setTitle('안내 사항 확인')
    .setChoiceValues([
      '참가비 110만원(부가세 포함) 안내를 확인했습니다.',
      '교재·식사·음료 및 다과 제공 안내를 확인했습니다.',
      '신청 후 등록 및 결제 안내가 별도로 진행된다는 점을 확인했습니다.'
    ])
    .setRequired(true);

  agreement.setValidation(
    FormApp.createCheckboxValidation()
      .requireSelectExactly(3)
      .build()
  );
}

function installFormSubmitTrigger_(form) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}

function onFormSubmit(e) {
  if (!e || !e.response) {
    throw new Error('폼 제출 이벤트 정보가 없습니다.');
  }

  const response = e.response;
  const respondentEmail = response.getRespondentEmail();
  const answers = mapResponses_(response.getItemResponses());

  sendAdminNotification_(respondentEmail, answers);
  if (respondentEmail) {
    sendApplicantConfirmation_(respondentEmail, answers);
  }
}

function mapResponses_(itemResponses) {
  const output = {};

  itemResponses.forEach((itemResponse) => {
    const title = itemResponse.getItem().getTitle();
    const value = itemResponse.getResponse();
    output[title] = Array.isArray(value) ? value.join(', ') : String(value);
  });

  return output;
}

function sendAdminNotification_(respondentEmail, answers) {
  const body = [
    '[주력 강의 재설계 과정] 신규 신청이 접수되었습니다.',
    '',
    '신청자 이메일: ' + (respondentEmail || '미수집'),
    '이름: ' + getValue_(answers, '이름'),
    '소속: ' + getValue_(answers, '소속'),
    '직무 또는 역할: ' + getValue_(answers, '직무 또는 역할'),
    '연락처: ' + getValue_(answers, '연락처'),
    '주력 강의명: ' + getValue_(answers, '주력 강의명'),
    '강의 주제 또는 핵심 내용: ' + getValue_(answers, '강의 주제 또는 핵심 내용'),
    '주요 학습자 또는 청중: ' + getValue_(answers, '주요 학습자 또는 청중'),
    '현재 강의 운영 상태: ' + getValue_(answers, '현재 강의 운영 상태'),
    '우선 개선 항목: ' + getValue_(answers, '우선 개선 항목'),
    '현재 가장 먼저 해결하고 싶은 문제: ' + getValue_(answers, '현재 가장 먼저 해결하고 싶은 문제'),
    '제출 가능한 자료: ' + getValue_(answers, '제출 가능한 자료'),
    '사전 검토 시 참고할 사항: ' + getValue_(answers, '사전 검토 시 참고할 사항'),
    '추천인 성함: ' + getValue_(answers, '추천인 성함')
  ].join('\n');

  MailApp.sendEmail({
    to: CONFIG.adminEmail,
    subject: '[신청 접수] ' + getValue_(answers, '이름') + ' / ' + getValue_(answers, '주력 강의명'),
    body
  });
}

function sendApplicantConfirmation_(recipient, answers) {
  const body = [
    getValue_(answers, '이름') + '님, 신청이 접수되었습니다.',
    '',
    CONFIG.courseTitle + ' 신청서를 제출해 주셔서 감사합니다.',
    '신청 내용 확인 후 등록 및 결제 안내를 순차적으로 드리겠습니다.',
    '',
    '신청 강의: ' + getValue_(answers, '주력 강의명'),
    '현재 강의 운영 상태: ' + getValue_(answers, '현재 강의 운영 상태'),
    '우선 개선 항목: ' + getValue_(answers, '우선 개선 항목'),
    '',
    '문의가 필요한 경우 회신해 주세요.'
  ].join('\n');

  MailApp.sendEmail({
    to: recipient,
    subject: '[신청 접수 완료] ' + CONFIG.courseTitle,
    body
  });
}

function formatResponseSheet_(sheet) {
  const firstSheet = sheet.getSheets()[0];
  firstSheet.setFrozenRows(1);
  firstSheet.autoResizeColumns(1, Math.max(firstSheet.getLastColumn(), 1));
}

function logProjectLinks_(form, sheet) {
  Logger.log('Form URL: ' + form.getPublishedUrl());
  Logger.log('Form Edit URL: ' + form.getEditUrl());
  Logger.log('Response Sheet URL: ' + sheet.getUrl());
}

function notifyProjectCreated_(form, sheet) {
  const body = [
    '[' + CONFIG.courseTitle + '] 신청서 생성이 완료되었습니다.',
    '',
    '신청서 URL: ' + form.getPublishedUrl(),
    '신청서 편집 URL: ' + form.getEditUrl(),
    '응답 시트 URL: ' + sheet.getUrl()
  ].join('\n');

  MailApp.sendEmail({
    to: CONFIG.adminEmail,
    subject: '[생성 완료] ' + CONFIG.formTitle,
    body
  });
}

function getProjectLinks() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Logger.log(JSON.stringify(props, null, 2));
  return props;
}

function getValue_(answers, key) {
  return answers[key] || '-';
}
