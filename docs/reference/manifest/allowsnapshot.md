---
title: Элемент AllowSnapshot в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f1aced0ce37b01c277ea5a8621f6c7764d2f761b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432349"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="73618-102">Элемент AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="73618-102">AllowSnapshot element</span></span>

<span data-ttu-id="73618-103">Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.</span><span class="sxs-lookup"><span data-stu-id="73618-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="73618-104">**Тип надстройки:** контентная</span><span class="sxs-lookup"><span data-stu-id="73618-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="73618-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="73618-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="73618-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="73618-106">Contained in</span></span>

[<span data-ttu-id="73618-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="73618-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="73618-108">Примечания</span><span class="sxs-lookup"><span data-stu-id="73618-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="73618-109">По умолчанию элементу **AllowSnapshot** присвоено значение `true`.</span><span class="sxs-lookup"><span data-stu-id="73618-109">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="73618-110">Это означает, что пользователи увидят изображение надстройки, если откроют документ в той версии ведущего приложения, которая не поддерживает надстройки Office. Кроме того, если ведущему приложению не удастся подключиться к серверу, на котором размещена надстройка, то отобразится статическое изображение надстройки.</span><span class="sxs-lookup"><span data-stu-id="73618-110">Security Note:AllowSnapshot is true by default. This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span> <span data-ttu-id="73618-111">Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="73618-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

