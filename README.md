# vbscript
itgen 플랫폼의 화면개발의 기본인 vbscript 의 import 시스템입니다.

최상단에 
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(import.vbs",1).readAll()
와 같은 코드를 넣어야합니다.

import.vbs는 import "package" 와 같은 형식의 code를 해석합니다.
또한, 다중 vbscript의 특성상 동일한 namespace의 package의 중복을 방지합니다.


예시)
'have to be in first line
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(import.vbs",1).readAll()
'data structure package
import "ds" 
