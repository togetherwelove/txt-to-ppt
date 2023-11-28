# txt-to-ppt
### 텍스트 파일의 가사를 개별 PowerPoint 슬라이드로 변환하는 VBA 매크로입니다.

해당 매크로는 텍스트 파일에 있는 찬양 가사를 읽어와서 각 가사를 PPT 슬라이드로 만드는 기능을 수행합니다.

매크로의 사용 방법은 다음과 같습니다:
1. 매크로를 실행하기 전에 찬양 가사가 저장된 텍스트 파일을 준비합니다.
2. F5를 눌러 슬라이드 쇼를 시작한 다음 **"매크로 실행"** 버튼을 누릅니다.
3. 매크로가 실행되면 파일 선택 대화상자가 표시됩니다. 여기에서 가사가 저장된 텍스트 파일을 선택합니다.
4. **"열기"** 버튼을 누르면 매크로가 해당 파일을 읽어들입니다. 파일은 **UTF-8**로 인코딩되어 있어야 합니다.
5. 각 가사는 별도의 PPT 슬라이드로 만들어집니다.
6. 모든 가사를 슬라이드로 만든 후, PPT 파일을 저장합니다.

파일 이름은 원본 텍스트 파일의 이름과 같으며, 확장자는 .pptx입니다.

---

### 아래는 실제 실행 화면입니다.

가사가 담긴 txt 파일입니다.

![image](https://github.com/togetherwelove/txt-to-ppt/assets/70801530/69164cde-a2a4-4cee-b73c-4463d5e78217)

한 슬라이드 내에서 개행을 표현하는 기호는 "/"입니다.
빈 칸에 "/"만 있다면 아무것도 없는 빈 슬라이드가 생성될 것입니다.

![image](https://github.com/togetherwelove/txt-to-ppt/assets/70801530/3d4b783a-e8bd-4995-b8d0-6ae104b84c2c)

"매크로 실행" 버튼을 누릅니다.

![image](https://github.com/togetherwelove/txt-to-ppt/assets/70801530/3a82f2ec-908f-4853-8299-ec461f75e495)

텍스트 파일을 선택하고 "열기" 버튼을 누릅니다.

![image](https://github.com/togetherwelove/txt-to-ppt/assets/70801530/c38e0902-be60-4e5c-9bb2-2075d71ad1e1)

성공적으로 매크로가 작동되었다면 이와 같은 메시지가 뜨게 됩니다.

![image](https://github.com/togetherwelove/txt-to-ppt/assets/70801530/90063d0a-1aa7-4db5-8c98-edbbd48ecc93)

텍스트 파일과 같은 경로에 같은 제목으로 .pptx 파일이 생성 되었습니다.

![image](https://github.com/togetherwelove/txt-to-ppt/assets/70801530/787dad8f-32b4-4f4d-8eed-ea4dd8f52efb)

열어보면 정상적으로 슬라이드가 생성된 것을 확인할 수 있습니다.

---

### 기본 세팅:

```
배경: 첫 번째 디자인의 첫 번째 레이아웃 (검은색 배경)
텍스트 위치: 가운데 위
텍스트 정렬: 가운데 정렬
폰트 크기: 28
폰트: HY견명조 (영어와 한글을 개별로 설정해야 합니다.)
폰트 색상: 흰색 [RGB(255, 255, 255)]
볼드 비활성화
텍스트 그림자 추가
```

기본 세팅을 바꾸시려면 VBA 매크로를 편집하여야 합니다.

```
    With myShape
        sentence(i) = Replace(sentence(i), "/", vbNewLine)
        With .TextFrame
            .TextRange.ParagraphFormat.Alignment = ppAlignCenter
            .TextRange.Text = sentence(i)

            ' 텍스트 프레임의 상단 여백을 설정합니다.
            .MarginTop = 0
            With .TextRange.Font

                ' 폰트 크기를 설정합니다.
                .Size = 28

                ' 영어의 폰트를 'HY견명조'로 설정합니다.
                .Name = "HY견명조"

                ' 한글의 폰트를 'HY견명조'로 설정합니다.
                .NameFarEast = "HY견명조"

                ' 폰트 색상을 흰색으로 설정합니다.
                .Color.RGB = RGB(255, 255, 255)

                ' 볼드체를 비활성화 합니다.
                .Bold = False

                ' 텍스트에 그림자를 추가합니다.
                .Shadow = True
            End With
        End With
    End With
```
