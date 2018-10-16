# <a name="allowsnapshot-element"></a><span data-ttu-id="b427e-101">Элемент AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="b427e-101">AllowSnapshot element</span></span>

<span data-ttu-id="b427e-102">Указывает, сохраняется ли моментальный снимок надстройки содержимого в документе ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="b427e-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="b427e-103">**Тип надстройки:** содержимое.</span><span class="sxs-lookup"><span data-stu-id="b427e-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="b427e-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b427e-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="b427e-105">Элемент, в котором содержится</span><span class="sxs-lookup"><span data-stu-id="b427e-105">Contained in:</span></span>

[<span data-ttu-id="b427e-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b427e-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b427e-107">Замечания</span><span class="sxs-lookup"><span data-stu-id="b427e-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="b427e-108">По умолчанию элементу **AllowSnapshot** задано значение `true`.</span><span class="sxs-lookup"><span data-stu-id="b427e-108">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="b427e-109">Это означает, что пользователи увидят изображение надстройки, если откроют документ в той версии ведущего приложения, которая не поддерживает надстройки Office. Кроме того, если ведущему приложению не удастся подключиться к серверу, на котором размещена надстройка, то отобразится статическое изображение надстройки.</span><span class="sxs-lookup"><span data-stu-id="b427e-109">Security Note:AllowSnapshot is true by default. This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span> <span data-ttu-id="b427e-110">Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="b427e-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

