---
title: Обеспечение совместимости надстройки Office с существующей надстройкой COM
description: Обеспечение совместимости с эквивалентной надстройкой COM, имеющей те же функциональные возможности, что и надстройка Office
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356911"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="9741e-103">Обеспечение совместимости надстройки Office с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="9741e-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="9741e-104">Если у вас есть надстройка COM, вы можете создать эквивалентную функцию в надстройке Office, чтобы расширить функциональные возможности решения на другие платформы, такие как Online или macOS.</span><span class="sxs-lookup"><span data-stu-id="9741e-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="9741e-105">Однако у надстроек Office нет всех функциональных возможностей, доступных в надстройках COM. Надстройка COM может улучшить работу, чем надстройка Office в Windows в Excel, Word и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="9741e-105">However, Office Add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Office Add-in on Windows in Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="9741e-106">Вы можете настроить надстройку Office таким образом, чтобы если на компьютере пользователя уже установлена эквивалентная надстройка COM, Office запускает надстройку COM, а не надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-106">You can configure your Office Add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Office Add-in.</span></span> <span data-ttu-id="9741e-107">Надстройка COM называется "эквивалентной", так как Office будет беспрепятственно переходить между надстройкой COM и надстройкой Office в зависимости от того, какая установлена в Windows.</span><span class="sxs-lookup"><span data-stu-id="9741e-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="9741e-108">Указание эквивалентной надстройки COM в манифесте</span><span class="sxs-lookup"><span data-stu-id="9741e-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="9741e-109">Чтобы обеспечить совместимость с существующей надстройкой COM, определите эквивалентную надстройку COM в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Office Add-in.</span></span> <span data-ttu-id="9741e-110">После этого Office будет использовать надстройку COM вместо надстройки Office при работе в Windows.</span><span class="sxs-lookup"><span data-stu-id="9741e-110">Then Office will use the COM add-in instead of your Office Add-in when running on Windows.</span></span>

<span data-ttu-id="9741e-111">`ProgID` Укажите эквивалентную надстройку COM.</span><span class="sxs-lookup"><span data-stu-id="9741e-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="9741e-112">Затем Office будет использовать пользовательский интерфейс надстройки COM, а не пользовательский интерфейс надстройки Office при установке надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="9741e-112">Office will then use the COM add-in UI instead of your Office Add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="9741e-113">В приведенном ниже примере показано, как указать как надстройку COM, так и XLL в качестве эквивалента.</span><span class="sxs-lookup"><span data-stu-id="9741e-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="9741e-114">Как правило, для полноты указывается и то, и другое, в этом примере показана как в контексте.</span><span class="sxs-lookup"><span data-stu-id="9741e-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="9741e-115">Они определяются по их `ProgID` и `FileName` соответственно.</span><span class="sxs-lookup"><span data-stu-id="9741e-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="9741e-116">Для получения дополнительной информации о совместимости XLL, ознакомьтесь [со статьЕй создание пользовательских функций, совместимых с пользовательскими ФУНКЦИЯМИ XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="9741e-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="9741e-117">Эквивалентное поведение для пользователей</span><span class="sxs-lookup"><span data-stu-id="9741e-117">Equivalent behavior for users</span></span>

<span data-ttu-id="9741e-118">Если в манифесте надстройки Office указана эквивалентная надстройка COM, Office подавляет пользовательский интерфейс надстройки Office в Windows при установке эквивалентной надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="9741e-118">When an equivalent COM add-in is specified in the Office Add-in manifest, Office suppresses your Office Add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="9741e-119">Это не повлияет на пользовательский интерфейс надстройки Office на других платформах, таких как Online или macOS.</span><span class="sxs-lookup"><span data-stu-id="9741e-119">This does not affect your Office Add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="9741e-120">Office скрывает только кнопки ленты и не запрещает установку.</span><span class="sxs-lookup"><span data-stu-id="9741e-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="9741e-121">Поэтому надстройка Office будет по-прежнему отображаться в следующих расположениях ПОЛЬЗОВАТЕЛЬСКОГО интерфейса:</span><span class="sxs-lookup"><span data-stu-id="9741e-121">Therefore your Office Add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="9741e-122">В разделе **Мои** надстройки, так как она технически устанавливается.</span><span class="sxs-lookup"><span data-stu-id="9741e-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="9741e-123">В качестве записи в диспетчере ленты.</span><span class="sxs-lookup"><span data-stu-id="9741e-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="9741e-124">В следующих сценариях описывается, что происходит в зависимости от того, как пользователь приобретает надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-124">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="9741e-125">AppSource приобретение надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9741e-125">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="9741e-126">Если пользователь загружает надстройку Office из AppSource, а аналогичная надстройка COM уже установлена, Office выполнит следующие действия:</span><span class="sxs-lookup"><span data-stu-id="9741e-126">If a user downloads the Office Add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="9741e-127">Установите надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-127">Install the Office Add-in.</span></span>
2. <span data-ttu-id="9741e-128">Скрытие ПОЛЬЗОВАТЕЛЬСКОГО интерфейса надстройки Office на ленте.</span><span class="sxs-lookup"><span data-stu-id="9741e-128">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="9741e-129">Отображение вызываемого абонента для пользователя, который указывает на кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="9741e-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="9741e-130">Централизованное развертывание надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9741e-130">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="9741e-131">Если Администратор развертывает надстройку Office в своем клиенте с помощью централизованного развертывания, а эквивалентная надстройка COM уже установлена, то пользователь должен перезапустить Office, прежде чем будут видны изменения.</span><span class="sxs-lookup"><span data-stu-id="9741e-131">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="9741e-132">После перезапуска Office будет:</span><span class="sxs-lookup"><span data-stu-id="9741e-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="9741e-133">Установите надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-133">Install the Office Add-in.</span></span>
2. <span data-ttu-id="9741e-134">Скрытие ПОЛЬЗОВАТЕЛЬСКОГО интерфейса надстройки Office на ленте.</span><span class="sxs-lookup"><span data-stu-id="9741e-134">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="9741e-135">Отображение вызываемого абонента для пользователя, который указывает на кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="9741e-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="9741e-136">Общий доступ к документу с помощью встроенной надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9741e-136">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="9741e-137">Если у пользователя установлена надстройка COM, а затем он получает общий документ с внедренной надстройкой Office, то при открытии документа Office будет:</span><span class="sxs-lookup"><span data-stu-id="9741e-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="9741e-138">ПредЛожит пользователю доверять надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-138">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="9741e-139">Если вы доверяете, надстройка Office будет установлена.</span><span class="sxs-lookup"><span data-stu-id="9741e-139">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="9741e-140">Скрытие ПОЛЬЗОВАТЕЛЬСКОГО интерфейса надстройки Office на ленте.</span><span class="sxs-lookup"><span data-stu-id="9741e-140">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="9741e-141">Другое поведение надстройки COM</span><span class="sxs-lookup"><span data-stu-id="9741e-141">Other COM add-in behavior</span></span>

<span data-ttu-id="9741e-142">Если пользователь удаляет надстройку COM, Office восстанавливает пользовательский интерфейс надстройки Office в Windows для эквивалентной установленной надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-142">If a user uninstalls the COM add-in, then Office restores the Office Add-in UI on Windows for the equivalent installed Office Add-in.</span></span>

<span data-ttu-id="9741e-143">Когда вы укажете эквивалентную надстройку COM для надстройки Office, Office прекратит обработку обновлений для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-143">Once you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="9741e-144">Пользователь должен удалить порядок надстроек COM для получения последних обновлений для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9741e-144">The user must uninstall the COM add-in order to get the latest updates for the Office Add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="9741e-145">См. также</span><span class="sxs-lookup"><span data-stu-id="9741e-145">See also</span></span>

- [<span data-ttu-id="9741e-146">Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL</span><span class="sxs-lookup"><span data-stu-id="9741e-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
