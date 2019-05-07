---
title: Обеспечение совместимости надстройки Excel с существующей надстройкой COM
description: Обеспечение совместимости с эквивалентной надстройкой COM, имеющей те же функциональные возможности, что и надстройка Excel
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628174"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="55a57-103">Обеспечение совместимости надстройки Office с существующей надстройкой COM (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="55a57-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="55a57-104">Если у вас есть надстройка COM, вы можете создать эквивалентную функцию в надстройке Excel, чтобы расширить функциональные возможности решения на другие платформы, такие как Online или macOS.</span><span class="sxs-lookup"><span data-stu-id="55a57-104">If you have an existing COM add-in, you can build equivalent functionality in your Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="55a57-105">Тем не менее, надстройки Excel не имеют всех функциональных возможностей, доступных в надстройках COM. Надстройка COM может обеспечить лучшую работу, чем надстройка Excel в Windows.</span><span class="sxs-lookup"><span data-stu-id="55a57-105">However, Excel add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Excel add-in on Windows.</span></span>

<span data-ttu-id="55a57-106">Вы можете настроить надстройку Excel таким образом, чтобы если на компьютере пользователя уже установлена эквивалентная надстройка COM, Office запускает надстройку COM, а не надстройку Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-106">You can configure your Excel add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Excel add-in.</span></span> <span data-ttu-id="55a57-107">Надстройка COM называется "эквивалентной", так как Office будет беспрепятственно переходить между надстройкой COM и надстройкой Excel в зависимости от того, какая установлена в Windows.</span><span class="sxs-lookup"><span data-stu-id="55a57-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Excel add-in depending on which is installed on Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="55a57-108">Указание эквивалентной надстройки COM в манифесте</span><span class="sxs-lookup"><span data-stu-id="55a57-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="55a57-109">Чтобы обеспечить совместимость с существующей надстройкой COM, определите эквивалентную надстройку COM в манифесте надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Excel add-in.</span></span> <span data-ttu-id="55a57-110">После этого Office будет использовать надстройку COM вместо надстройки Excel при работе в Windows.</span><span class="sxs-lookup"><span data-stu-id="55a57-110">Then Office will use the COM add-in instead of your Excel add-in when running on Windows.</span></span>

<span data-ttu-id="55a57-111">`ProgID` Укажите эквивалентную надстройку COM.</span><span class="sxs-lookup"><span data-stu-id="55a57-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="55a57-112">Затем Office будет использовать пользовательский интерфейс надстройки COM, а не пользовательский интерфейс надстройки Excel при установке надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="55a57-112">Office will then use the COM add-in UI instead of your Excel add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="55a57-113">В приведенном ниже примере показано, как указать как надстройку COM, так и XLL в качестве эквивалента.</span><span class="sxs-lookup"><span data-stu-id="55a57-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="55a57-114">Как правило, для полноты указывается и то, и другое, в этом примере показана как в контексте.</span><span class="sxs-lookup"><span data-stu-id="55a57-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="55a57-115">Они определяются по их `ProgID` и `FileName` соответственно.</span><span class="sxs-lookup"><span data-stu-id="55a57-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="55a57-116">Для получения дополнительной информации о совместимости XLL, ознакомьтесь [со статьей Создание пользовательских функций, совместимых с пользовательскими ФУНКЦИЯМИ XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="55a57-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

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

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="55a57-117">Эквивалентное поведение для пользователей</span><span class="sxs-lookup"><span data-stu-id="55a57-117">Equivalent behavior for users</span></span>

<span data-ttu-id="55a57-118">Если в манифесте надстройки Excel указана эквивалентная надстройка COM, Office отключает пользовательский интерфейс надстройки Excel в Windows при установке эквивалентной надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="55a57-118">When an equivalent COM add-in is specified in the Excel add-in manifest, Office suppresses your Excel add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="55a57-119">Это не влияет на пользовательский интерфейс надстройки Excel на других платформах, таких как Online или macOS.</span><span class="sxs-lookup"><span data-stu-id="55a57-119">This does not affect your Excel add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="55a57-120">Office скрывает только кнопки ленты и не запрещает установку.</span><span class="sxs-lookup"><span data-stu-id="55a57-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="55a57-121">Поэтому надстройка Excel по-прежнему будет отображаться в следующих расположениях ПОЛЬЗОВАТЕЛЬСКОГО интерфейса:</span><span class="sxs-lookup"><span data-stu-id="55a57-121">Therefore your Excel add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="55a57-122">В разделе **Мои** надстройки, так как она технически устанавливается.</span><span class="sxs-lookup"><span data-stu-id="55a57-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="55a57-123">В качестве записи в диспетчере ленты.</span><span class="sxs-lookup"><span data-stu-id="55a57-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="55a57-124">В следующих сценариях описывается, что происходит в зависимости от того, как пользователь приобретает надстройку Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-124">The following scenarios describe what happens depending on how the user acquires the Excel add-in.</span></span>

### <a name="appsource-acquisition-of-an-excel-add-in"></a><span data-ttu-id="55a57-125">AppSource получение надстройки Excel</span><span class="sxs-lookup"><span data-stu-id="55a57-125">AppSource acquisition of an Excel add-in</span></span>

<span data-ttu-id="55a57-126">Если пользователь загружает надстройку Excel из AppSource, а аналогичная надстройка COM уже установлена, Office выполнит следующие действия:</span><span class="sxs-lookup"><span data-stu-id="55a57-126">If a user downloads the Excel add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="55a57-127">Установите надстройку Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-127">Install the Excel add-in.</span></span>
2. <span data-ttu-id="55a57-128">Скрыть пользовательский интерфейс надстройки Excel на ленте.</span><span class="sxs-lookup"><span data-stu-id="55a57-128">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="55a57-129">Отображение вызываемого абонента для пользователя, который указывает на кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="55a57-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-excel-add-in"></a><span data-ttu-id="55a57-130">Централизованное развертывание надстройки Excel</span><span class="sxs-lookup"><span data-stu-id="55a57-130">Centralized deployment of Excel add-in</span></span>

<span data-ttu-id="55a57-131">Если Администратор развертывает надстройку Excel в своем клиенте с помощью централизованного развертывания, а эквивалентная надстройка COM уже установлена, то пользователю необходимо перезапустить Office до того, как будут видны изменения.</span><span class="sxs-lookup"><span data-stu-id="55a57-131">If an admin deploys the Excel add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="55a57-132">После перезапуска Office будет:</span><span class="sxs-lookup"><span data-stu-id="55a57-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="55a57-133">Установите надстройку Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-133">Install the Excel add-in.</span></span>
2. <span data-ttu-id="55a57-134">Скрыть пользовательский интерфейс надстройки Excel на ленте.</span><span class="sxs-lookup"><span data-stu-id="55a57-134">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="55a57-135">Отображение вызываемого абонента для пользователя, который указывает на кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="55a57-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-excel-add-in"></a><span data-ttu-id="55a57-136">Общий доступ к документу с помощью встроенной надстройки Excel</span><span class="sxs-lookup"><span data-stu-id="55a57-136">Document shared with embedded Excel add-in</span></span>

<span data-ttu-id="55a57-137">Если у пользователя установлена надстройка COM, а затем он получает общий документ с внедренной надстройкой Excel, то при открытии документа Office будет:</span><span class="sxs-lookup"><span data-stu-id="55a57-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Excel add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="55a57-138">Предложит пользователю доверять надстройке Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-138">Prompt the user to trust the Excel add-in.</span></span>
2. <span data-ttu-id="55a57-139">Если вы доверяете, надстройка Excel будет установлена.</span><span class="sxs-lookup"><span data-stu-id="55a57-139">If trusted, the Excel add-in will install.</span></span>
3. <span data-ttu-id="55a57-140">Скрыть пользовательский интерфейс надстройки Excel на ленте.</span><span class="sxs-lookup"><span data-stu-id="55a57-140">Hide the Excel add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="55a57-141">Другое поведение надстройки COM</span><span class="sxs-lookup"><span data-stu-id="55a57-141">Other COM add-in behavior</span></span>

<span data-ttu-id="55a57-142">Если пользователь удаляет надстройку COM, Office восстанавливает пользовательский интерфейс надстройки Excel в Windows для эквивалентной установленной надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-142">If a user uninstalls the COM add-in, then Office restores the Excel add-in UI on Windows for the equivalent installed Excel add-in.</span></span>

<span data-ttu-id="55a57-143">Когда вы укажете эквивалентную надстройку COM для надстройки Excel, Office прекратит обработку обновлений для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-143">Once you specify an equivalent COM add-in for your Excel add-in, Office stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="55a57-144">Пользователь должен удалить порядок надстроек COM для получения последних обновлений для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="55a57-144">The user must uninstall the COM add-in order to get the latest updates for the Excel add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="55a57-145">См. также</span><span class="sxs-lookup"><span data-stu-id="55a57-145">See also</span></span>

- [<span data-ttu-id="55a57-146">Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL</span><span class="sxs-lookup"><span data-stu-id="55a57-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
