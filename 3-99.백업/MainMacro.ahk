;==========키별 특수문자==========
;Control : ^
;Alt     : !
;Shift   : +
;Win Key : #
;==========키별 특수문자==========

; ;엑셀 교재 필사용_S
; ^+Right::SendInput "^+{S}"
; ^+Left::SendInput "^+{A}"
; ^+!Right::SendInput "^+{W}"
; ^+!Left::SendInput "^+{Q}"
; +!Right::SendInput "^+{F}"
; +!Left::SendInput "^+{D}"
; ;CapsLock::SendInput "^+{T}"
; ;^+Del::SendInput "^+{Y}"
; ;^+End::SendInput "^+{O}"
; ^!a::SendInput "!{Enter}" ;;엑셀 편집모드에서 엔터(alt+Enter)
; ;엑셀 교재 필사용_E

; #HotIf WinActive("ahk_exe EXCEL.exe") ;EXCEL 내에서만 실행
; F8 & 1::SendInput "➊"
; F8 & 2::SendInput "➋"
; F8 & 3::SendInput "➌"
; F8 & 4::SendInput "➍"
; F8 & 5::SendInput "➎"
; F8 & 6::SendInput "➏"
; F8 & 7::SendInput "➐"
; F8 & 8::SendInput "➑"
; F8 & 9::SendInput "➒"
; F8 & 0::SendInput "➓"

;^Right::SendInput "  "
!`::SendInput "・"
!-::SendInput "－"
![::SendInput "‘"
!]::SendInput "’"
!;::SendInput "“"
!'::SendInput "”"
!,::SendInput "〈"
!.::SendInput "〉"
!9::SendInput "「"
!0::SendInput "」"
!a::SendInput " × → "
!s::SendInput "Ⅴ"
; #HotIf ;조건 해제 (이후 작성되는 다른 단축키는 글로벌 적용)

#HotIf WinActive("ahk_exe Code.exe") ;Visual Studio Code 내에서만 실행
F8 & 1::SendInput "➊"
F8 & 2::SendInput "➋"
F8 & 3::SendInput "➌"
F8 & 4::SendInput "➍"
F8 & 5::SendInput "➎"
F8 & 6::SendInput "➏"
F8 & 7::SendInput "➐"
F8 & 8::SendInput "➑"
F8 & 9::SendInput "➒"
F8 & 0::SendInput "➓"

F9 & 1::SendInput "Ⅰ."
F9 & 2::SendInput "Ⅱ."
F9 & 3::SendInput "Ⅲ."
F9 & 4::SendInput "Ⅳ."
F9 & 5::SendInput "Ⅴ."
F9 & 6::SendInput "Ⅵ."
F9 & 7::SendInput "Ⅶ."
F9 & 8::SendInput "Ⅷ."
F9 & 9::SendInput "Ⅸ."
F9 & 0::SendInput "Ⅹ."

F10 & 1::SendInput "1."
F10 & 2::SendInput "2."
F10 & 3::SendInput "3."
F10 & 4::SendInput "4."
F10 & 5::SendInput "5."
F10 & 6::SendInput "6."
F10 & 7::SendInput "7."
F10 & 8::SendInput "8."
F10 & 9::SendInput "9."
F10 & 0::SendInput "10."

F11 & 1::SendInput "(1)"
F11 & 2::SendInput "(2)"
F11 & 3::SendInput "(3)"
F11 & 4::SendInput "(4)"
F11 & 5::SendInput "(5)"
F11 & 6::SendInput "(6)"
F11 & 7::SendInput "(7)"
F11 & 8::SendInput "(8)"
F11 & 9::SendInput "(9)"
F11 & 0::SendInput "(10)"

F12 & 1::SendInput "①"
F12 & 2::SendInput "②"
F12 & 3::SendInput "③"
F12 & 4::SendInput "④"
F12 & 5::SendInput "⑤"
F12 & 6::SendInput "⑥"
F12 & 7::SendInput "⑦"
F12 & 8::SendInput "⑧"
F12 & 9::SendInput "⑨"
F12 & 0::SendInput "⑩"

; =========================================================
; 배경 강조 공통 함수 (시작문자, 종료문자 매개변수 전달)
; =========================================================
#Requires AutoHotkey v2.0

; 전역 변수 선언 (직전 실행 동작 저장용)
global g_LastAction := ""

; 래퍼 함수 : 실행할 함수와 인자를 받아 전역 변수에 저장 후 실행
ExecuteAction(fn, params*) 
{
    global g_LastAction
    if HasMethod(fn) 
    {
        g_LastAction := fn.Bind(params*)
        return fn(params*)
    }
}

; 직전 실행 함수 재호출
RepeatLastAction() 
{
    global g_LastAction
    if IsObject(g_LastAction) 
    {
        return g_LastAction()
    }
}

; 텍스트 감싸기 및 특수문자 정화 함수
WrapOrCleanSelectedText(openChar := "", closeChar := "") 
{
    ; 기존 클립보드 내용 백업
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    
    ; 현재 선택한 텍스트 복사
    Send("^c")
    
    ; 복사 대기 진행
    if !ClipWait(0.3)
    {
        A_Clipboard := ClipSaved
        ClipSaved := ""
        return
    }
    
    ; 복사된 텍스트 추출
    copiedText := A_Clipboard
    
    ; [오동작 방지] 선택 영역이 없거나 에디터의 행 전체 자동 복사가 동작한 경우 통합 차단
    if (copiedText = "") || RegExMatch(copiedText, "^\r?\n$") || (RegExMatch(copiedText, "[\r\n]") && RegExMatch(copiedText, "^[^\r\n]+\r?\n$"))
    {
        A_Clipboard := ClipSaved
        ClipSaved := ""
        return
    }

    ; 선택한 텍스트에서 지정된 특수문자(' " ‘’ “” 〈〉 「」) 모두 삭제
    cleanedText := RegExReplace(copiedText, "['`"‘’“”〈〉「」*]")
    
    ; 매개변수로 넘어온 시작/종료 문자로 감싸기
    A_Clipboard := openChar . cleanedText . closeChar
    
    ; 붙여넣기 실행
    Send("^v")
    
    ; 원래 클립보드 복원 처리 (클립보드 반영 안정성을 위해 대기 후 복원)
    Sleep(200)
    A_Clipboard := ClipSaved
    ClipSaved := ""
}

; 선택 영역 강조용 따옴표(' ‘ ’) 모두 삭제하는 독립 함수
DeleteQuotes()
{
    ; 기존 클립보드 내용 백업
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    
    ; 현재 선택한 텍스트 복사
    Send("^c")
    
    ; 복사 대기 진행
    if !ClipWait(0.3)
    {
        A_Clipboard := ClipSaved
        ClipSaved := ""
        return
    }
    
    ; 복사된 텍스트 추출
    copiedText := A_Clipboard
    
    ; [오동작 방지] 선택 영역이 없거나 에디터의 행 전체 자동 복사가 동작한 경우 통합 차단
    if (copiedText = "") || RegExMatch(copiedText, "^\r?\n$") || (RegExMatch(copiedText, "[\r\n]") && RegExMatch(copiedText, "^[^\r\n]+\r?\n$")) 
    {
        A_Clipboard := ClipSaved
        ClipSaved := ""
        return
    }
    
    ; 선택한 영역의 강조용 따옴표(' ‘ ’) 모두 삭제
    cleanedText := RegExReplace(copiedText, "['‘’]")
    
    ; 변환된 텍스트 클립보드 설정
    A_Clipboard := cleanedText
    
    ; 붙여넣기 실행
    Send("^v")
    
    ; 원래 클립보드 복원 처리 (클립보드 반영 안정성을 위해 대기 후 복원)
    Sleep(200)
    A_Clipboard := ClipSaved
    ClipSaved := ""
}

; 기존 ' 삭제, ‘’를 '로 교체하는 독립 함수
ReplaceSingleQuotes() 
{
    ; 기존 클립보드 내용 백업
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    
    ; 현재 선택한 텍스트 복사
    Send("^c")
    
    ; 복사 대기 진행
    if !ClipWait(0.3)
    {
        A_Clipboard := ClipSaved
        ClipSaved := ""
        return
    }
    
    ; 복사된 텍스트 추출
    copiedText := A_Clipboard
    
    ; [오동작 방지] 선택 영역이 없거나 에디터의 행 전체 자동 복사가 동작한 경우 통합 차단
    if (copiedText = "") || RegExMatch(copiedText, "^\r?\n$") || (RegExMatch(copiedText, "[\r\n]") && RegExMatch(copiedText, "^[^\r\n]+\r?\n$")) 
    {
        A_Clipboard := ClipSaved
        ClipSaved := ""
        return
    }
    
    ; 1. 선택한 영역의 작은따옴표(') 모두 삭제
    cleanedText := RegExReplace(copiedText, "'", "")
    
    ; 2. 선택한 영역의 홑따옴표(‘ 및 ’)를 작은따옴표(')로 교체
    cleanedText := RegExReplace(cleanedText, "[‘’]", "'")
    
    ; 변환된 텍스트 클립보드 설정
    A_Clipboard := cleanedText
    
    ; 붙여넣기 실행
    Send("^v")
    
    ; 원래 클립보드 복원 처리 (클립보드 반영 안정성을 위해 대기 후 복원)
    Sleep(200)
    A_Clipboard := ClipSaved
    ClipSaved := ""
}

; 단축키 매핑 (ExecuteAction을 거쳐 실행 이력 기록)
!1::ExecuteAction(WrapOrCleanSelectedText, "'", "'") ;기본 강조
!2::ExecuteAction(WrapOrCleanSelectedText, "‘", "’") ;키워드 강조
!3::ExecuteAction(WrapOrCleanSelectedText, "`"", "`"") ;법령, 학자이름
!4::ExecuteAction(WrapOrCleanSelectedText, "「", "」") ;법령 전체
!5::ExecuteAction(WrapOrCleanSelectedText, "“", "”") ;판례, 학자저서(논문)
!6::ExecuteAction(WrapOrCleanSelectedText, "〈", "〉") ;쟁점
!d::ExecuteAction(WrapOrCleanSelectedText) ;전체 특수문자 삭제
!c::ExecuteAction(DeleteQuotes) ;강조용 따옴표 전체 삭제
!r::ExecuteAction(ReplaceSingleQuotes) ;강조용 따옴표 대체(‘’ → ')

; F4 키 : 직전 수행 기능 재실행
F4::RepeatLastAction()

; Ⅰ. 오토핫키 직전 실행 기능 구현 구조 및 원리
;   1. 함수 객체(BoundFunc)와 전역 상태 저장 메커니즘
;     (1) BoundFunc 객체를 통한 매개변수 고정 및 상태 보존
;      －오토핫키 v2 환경에서 특정 함수와 해당 함수에 전달된 매개변수(인자)를 차후에 동일하게 다시 실행하려면, 실행 시점의 함수 포인터와 매개변수를 하나로 묶어 메모리에 보존해야 한다. 이를 지원하는 핵심 기능이 "Bind" 메서드이다. "Bind" 메서드를 사용하면 함수에 특정 매개변수가 미리 결합된 "BoundFunc" 객체가 생성된다. 해당 객체는 전역 변수에 저장되어 향후 인자 전달 없이 호출하더라도 최초 바인딩된 매개변수를 그대로 유지하며 실행된다.
;     (2) 래퍼(Wrapper) 함수를 이용한 중앙 실행 추적
;      －단축키(!1 ~ !8)가 눌릴 때마다 실행 대상 함수를 직접 호출하지 않고, 중간에서 실행 정보를 기록하는 래퍼 함수인 "ExecuteAction"을 거치도록 설계한다. "ExecuteAction" 함수는 전달받은 함수와 매개변수를 바탕으로 "BoundFunc" 객체를 만들어 전역 변수 "g_LastAction"에 할당하고, 그 즉시 해당 함수를 실행한다. 이를 통해 어떤 단축키가 실행되더라도 항상 가장 최근에 수행된 기능이 전역 변수에 갱신된다.
;   2. 익명 함수 구조 개편 필요성 : 익명 함수의 한계와 독자적 함수 분리
;     －제시된 기존 코드에서 "!8" 단축키는 블록 형태의 익명 함수로 구현되어 있다. 익명 함수는 외부에서 함수 이름을 직접 참조하거나 바인딩하여 전역 변수에 저장하기에 구조적인 제약이 따른다. 따라서 "!8" 단축키의 내부 로직을 "ReplaceSingleQuotes"라는 독립된 이름의 함수로 분리하여 정의한다. 이를 통해 "!1"부터 "!7"까지 사용된 "WrapOrCleanSelectedText" 함수와 동일하게 래퍼 함수를 통한 바인딩 및 저장 처리가 가능해진다.

; Ⅱ. 직전 실행 기능(F4) 보완 전체 오토핫키 스크립트
;   1. 완전한 수정 소스코드 : 전체 오토핫키 v2 코드
;     －아래 소스코드는 기존 작성된 로직의 안정성을 유지하면서, F4 키 입력 시 직전에 실행했던 텍스트 변환 및 감싸기 기능을 동일하게 수행하도록 보완한 완성형 스크립트이다.

; Ⅲ. 소스코드 핵심 변경 사항 및 구체적 설명
;   1. 전역 변수 및 래퍼 함수 도입
;     (1) g_LastAction 전역 변수 선언
;      －스크립트 최상단에 "global g_LastAction := "을 선언하여 가장 최근에 호출된 함수 객체를 저장할 수 있는 공간을 마련한다. 오토핫키 v2에서는 변수의 스코프 관리가 엄격하므로 전역 변수를 명시적으로 선언하는 것이 안전하다.
;     (2) ExecuteAction 래퍼 함수 동작 원리
;      －"ExecuteAction(fn, params*)" 함수는 단축키 입력 시 실제 동작을 중계한다. "params*" 가변 인자 구문을 통해 전달받은 인자들을 "fn.Bind(params*)"로 결합하여 "BoundFunc" 객체를 생성한다. 해당 객체는 "g_LastAction"에 보존되며, 동시에 "fn(params*)"를 실행하여 사용자가 요청한 원래의 동작을 수행한다.
;     (3) RepeatLastAction 및 F4 단축키 바인딩
;      －"RepeatLastAction" 함수는 "IsObject(g_LastAction)" 구문을 통해 "g_LastAction"에 유효한 함수 객체가 저장되어 있는지 확인한다. 저장된 객체가 존재하는 경우 "g_LastAction()"을 호출하여 직전 매개변수가 포함된 함수를 즉시 재실행한다. 이를 "F4::RepeatLastAction()" 단축키에 연결하여 원터치 재실행 기능을 완성한다.
;   2. !8 익명 함수의 ReplaceSingleQuotes 분리 : 익명 함수에서 독립 함수로의 구조적 전환
;     －기존 "!8:: {" 형태의 블록 구현 방식은 단축키 바인딩과 로직이 결합되어 있어 다른 단축키에서 해당 로직을 재호출하기 어려웠다. 이를 "ReplaceSingleQuotes()"라는 이름의 독립 함수로 추출하여 작성함으로써 "!8::ExecuteAction(ReplaceSingleQuotes)" 형태로 깔끔하게 등록할 수 있게 되었다. 이로 인해 F4 키를 눌렀을 때 "!8"의 동작 역시 완벽하게 재실행된다.
;   3. 예외 처리 및 오동작 방지 메커니즘
;     (1) 최초 실행 시 g_LastAction 유효성 검사
;      －스크립트를 실행한 직후, 사용자가 "!1" ~ "!8" 중 어떠한 단축키도 누르지 않은 상태에서 F4 키를 누를 수 있다. 이때 "IsObject(g_LastAction)" 구문이 거짓(False)을 반환하므로 런타임 오류나 예외가 발생하지 않고 안전하게 무시된다.
;     (2) 클립보드 복원 대기 및 메모리 관리
;      －기존 코드에 포함되어 있던 "ClipWait(0.3)" 복사 대기 구문과 "Sleep(200)" 클립보드 복원 대기 구문은 그대로 유지된다. F4 키를 통해 재실행될 때에도 동일한 클립보드 백업 및 복원 로직이 수행되므로 타겟 에디터에서의 텍스트 교체 작업이 안정적으로 이루어진다.

; Ⅳ. 결론
;   －제시된 기존 오토핫키 v2 스크립트에 F4 키를 통한 직전 기능 재실행 로직을 성공적으로 이식하였다. 핵심은 "ExecuteAction" 래퍼 함수와 "Bind" 메서드를 도입하여 호출된 함수와 인자를 "g_LastAction" 전역 변수에 보존하는 것에 있다. 또한 블록 형태로 작성되어 있던 "!8" 단축키 로직을 "ReplaceSingleQuotes" 독립 함수로 분리하여 모든 단축키 기능이 동일한 재실행 매커니즘을 공유하도록 개편하였다. 보완된 소스코드를 적용하여 사용할 경우 작업 효율성을 극대화할 수 있다.
#HotIf ;조건 해제 (이후 작성되는 다른 단축키는 글로벌 적용)