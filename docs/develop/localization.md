---
title: Локализация надстроек для Office
description: Используйте API Office JavaScript для определения локального кода и отображения строк на основе Office приложения, а также для интерпретации или отображения данных на основе локального кода данных.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: b49d64f2c9391539ac2d5929ebff2a4ecc08b630
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349828"
---
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="44afe-103">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="44afe-103">Localization for Office Add-ins</span></span>

<span data-ttu-id="44afe-104">Вы можете реализовать любую схему локализации, которая подходит вашему Надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="44afe-104">You can implement any localization scheme that's appropriate for your Office Add-in.</span></span> <span data-ttu-id="44afe-105">API JavaScript и схема манифеста платформы Надстройки Office предоставляют несколько вариантов.</span><span class="sxs-lookup"><span data-stu-id="44afe-105">The JavaScript API and manifest schema of the Office Add-ins platform provide some choices.</span></span> <span data-ttu-id="44afe-106">Вы можете использовать Office API JavaScript для определения локального кода и отображения строк на основе Office приложения, а также для интерпретации или отображения данных на основе локального кода данных.</span><span class="sxs-lookup"><span data-stu-id="44afe-106">You can use the Office JavaScript API to determine a locale and display strings based on the locale of the Office application, or to interpret or display data based on the locale of the data.</span></span> <span data-ttu-id="44afe-107">Вы можете использовать манифест, чтобы указать расположение файла надстройки и описательной информации, зависящих от языковых параметров.</span><span class="sxs-lookup"><span data-stu-id="44afe-107">You can use the manifest to specify locale-specific add-in file location and descriptive information.</span></span> <span data-ttu-id="44afe-108">Либо можно использовать сценарий Microsoft Ajax для поддержки глобализации и локализации.</span><span class="sxs-lookup"><span data-stu-id="44afe-108">Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="44afe-109">Определение параметров, зависящих от языка, с помощью API JavaScript</span><span class="sxs-lookup"><span data-stu-id="44afe-109">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="44afe-110">API Office JavaScript предоставляет два свойства, поддерживаюющие отображение или интерпретацию значений, совместимых с локальным Office приложения и данных:</span><span class="sxs-lookup"><span data-stu-id="44afe-110">The Office JavaScript API provides two properties that support displaying or interpreting values consistent with the locale of the Office application and data:</span></span>

- <span data-ttu-id="44afe-111">[DisplayLanguage Context.displayLanguage][] указывает локализ (или язык) пользовательского интерфейса Office приложения.</span><span class="sxs-lookup"><span data-stu-id="44afe-111">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the Office application.</span></span> <span data-ttu-id="44afe-112">В следующем примере проверяется, Office приложение использует локальный код en-US или fr-FR и отображает приветствие, определенное для локального.</span><span class="sxs-lookup"><span data-stu-id="44afe-112">The following example verifies if the Office application uses the en-US or fr-FR locale, and displays a locale-specific greeting.</span></span>

    ```js
    function sayHelloWithDisplayLanguage() {
        var myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
            case 'en-US':
                write('Hello!');
                break;
            case 'fr-FR':
                write('Bonjour!');
                break;
        }
    }

    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }
    ```

- <span data-ttu-id="44afe-p103">[Context.contentLanguage][contentLanguage] задает языковой стандарт данных. Вы можете не проверять свойство [displayLanguage], а назначить свойству [contentLanguage] значение `myLanguage` и воспользоваться тем же кодом для отображения приветствия на языке данных:</span><span class="sxs-lookup"><span data-stu-id="44afe-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` the value of the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="44afe-115">Управление локализацией через манифест</span><span class="sxs-lookup"><span data-stu-id="44afe-115">Control localization from the manifest</span></span>


<span data-ttu-id="44afe-116">Каждое Надстройка Office задает в своем манифесте элемент [DefaultLocale] и языковой параметр.</span><span class="sxs-lookup"><span data-stu-id="44afe-116">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest.</span></span> <span data-ttu-id="44afe-117">По умолчанию Office и клиентские приложения Office применяют значения элементов [Description,] [DisplayName,] [IconUrl,] [HighResolutionIconUrl]и [SourceLocation.]</span><span class="sxs-lookup"><span data-stu-id="44afe-117">By default, the Office Add-in platform and Office client applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales.</span></span> <span data-ttu-id="44afe-118">Чтобы изменить значения для определенных языковых стандартов, укажите для любого из этих пяти элементов дочерний элемент [Override].</span><span class="sxs-lookup"><span data-stu-id="44afe-118">You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements.</span></span> <span data-ttu-id="44afe-119">Значение элемента [DefaultLocale] и атрибута `Locale` элемента [Override] указывается в соответствии со спецификацией [RFC 3066], "Теги для идентификации языков".</span><span class="sxs-lookup"><span data-stu-id="44afe-119">The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages."</span></span> <span data-ttu-id="44afe-120">В таблице 1 описана поддержка локализации для этих элементов.</span><span class="sxs-lookup"><span data-stu-id="44afe-120">Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="44afe-121">*Таблица 1. Поддержка локализации*</span><span class="sxs-lookup"><span data-stu-id="44afe-121">*Table 1. Localization support*</span></span>


|<span data-ttu-id="44afe-122">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="44afe-122">**Element**</span></span>|<span data-ttu-id="44afe-123">**Поддержка локализации**</span><span class="sxs-lookup"><span data-stu-id="44afe-123">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="44afe-124">[Описание]</span><span class="sxs-lookup"><span data-stu-id="44afe-124">[Description]</span></span>   |<span data-ttu-id="44afe-125">Для каждого заданного языкового стандарта пользователи могут видеть локализованное описание надстройки в AppSource (или частном каталоге).</span><span class="sxs-lookup"><span data-stu-id="44afe-125">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="44afe-126">В случае надстроек Outlook пользователи смогут увидеть описание в Центре администрирования Exchange после установки.</span><span class="sxs-lookup"><span data-stu-id="44afe-126">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="44afe-127">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="44afe-127">[DisplayName]</span></span>   |<span data-ttu-id="44afe-128">Для каждого заданного языкового стандарта пользователи могут видеть локализованное описание надстройки в AppSource (или частном каталоге).</span><span class="sxs-lookup"><span data-stu-id="44afe-128">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="44afe-129">В случае надстроек Outlook пользователи смогут увидеть отображаемое имя в качестве метки для кнопки надстройки Outlook и в Центре администрирования Exchange после установки.</span><span class="sxs-lookup"><span data-stu-id="44afe-129">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="44afe-130">В случае контентных надстроек и надстроек области задач пользователи могут видеть отображаемое имя на ленте после установки надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-130">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="44afe-131">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="44afe-131">[IconUrl]</span></span>        |<span data-ttu-id="44afe-p105">Изображение значка является необязательным. Можно использовать ту же методику переопределений, чтобы задать определенное изображение для определенной культуры. Если вы используете значок и локализуете его, пользователи с заданными языковыми параметрами могут видеть локализованный значок надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="44afe-135">В случае надстроек Outlook пользователи могут видеть значок в Центре администрирования Exchange после установки надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-135">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="44afe-136">После установки надстроек области задач и контентных надстроек пользователи видят значок на ленте.</span><span class="sxs-lookup"><span data-stu-id="44afe-136">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="44afe-137">[HighResolutionIconUrl] **Важно!** Этот элемент доступен только для надстроек, использующих схему манифеста версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="44afe-137">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="44afe-p106">Изображение значка с высоким разрешением не является обязательным, но если оно указано, то должно находиться после элемента [IconUrl]. Если указан параметр [HighResolutionIconUrl] и надстройка установлена на устройстве, поддерживающем высокое разрешение, то вместо значения [IconUrl] используется значение [HighResolutionIconUrl].</span><span class="sxs-lookup"><span data-stu-id="44afe-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="44afe-p107">Можно использовать ту же методику переопределений, чтобы задать определенное изображение для определенной культуры. Если вы используете значок и локализуете его, пользователи с заданными языковыми параметрами могут видеть локализованный значок надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="44afe-142">В случае надстроек Outlook пользователи могут видеть значок в Центре администрирования Exchange после установки надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-142">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="44afe-143">После установки надстроек области задач и контентных надстроек пользователи видят значок на ленте.</span><span class="sxs-lookup"><span data-stu-id="44afe-143">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="44afe-144">[Resources] **Важно!** Этот элемент доступен только для надстроек, в которых используется схема манифеста версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="44afe-144">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="44afe-145">Для пользователей в каждой указываемой вами локали отображаются ресурсы строк и значков, которые вы специально создаете для надстройки в этой локали.</span><span class="sxs-lookup"><span data-stu-id="44afe-145">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="44afe-146">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="44afe-146">[SourceLocation]</span></span>   |<span data-ttu-id="44afe-147">Пользователи каждого языкового стандарта видят веб-страницу, специально разработанную для использования надстройки с этим стандартом.</span><span class="sxs-lookup"><span data-stu-id="44afe-147">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> [!NOTE]
> <span data-ttu-id="44afe-148">Локализовать описание и отображаемое имя можно только для языковых стандартов, которые поддерживаются в Office.</span><span class="sxs-lookup"><span data-stu-id="44afe-148">You can localize the description and display name for only the locales that Office supports.</span></span> <span data-ttu-id="44afe-149">Список языков и языковых стандартов для текущего выпуска Office см. в статье [Идентификаторы языков и значения OptionState Id в Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="44afe-149">See [Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="44afe-150">Примеры</span><span class="sxs-lookup"><span data-stu-id="44afe-150">Examples</span></span>

<span data-ttu-id="44afe-p109">Например, надстройка Office может задать для параметра [DefaultLocale] значения `en-us`. Для элемента [DisplayName] надстройка может задать дочерний элемент [Override], соответствующий языковому стандарту `fr-fr`, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="44afe-p109">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`. For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span>


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> <span data-ttu-id="44afe-153">Если вам необходимо локализовать несколько областей в семействе языков, например `de-de` и `de-at`, рекомендуется разделить элементы `Override` для каждой области.</span><span class="sxs-lookup"><span data-stu-id="44afe-153">If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area.</span></span> <span data-ttu-id="44afe-154">Использование только только языкового имени в данном случае не поддерживается во всех Office клиентских приложений `de` и платформ.</span><span class="sxs-lookup"><span data-stu-id="44afe-154">Using just the language name alone, in this case, `de`, is not supported across all combinations of Office client applications and platforms.</span></span>

<span data-ttu-id="44afe-p111">Это значит, что по умолчанию надстройка использует языковой стандарт `en-us`. Пользователи видят отображаемое имя Video player (видеопроигрыватель) на английском языке для всех языковых стандартов за исключением случаев, когда на клиентском компьютере используется языковой стандарт `fr-fr`. В этом случае пользователи увидят отображаемое имя Lecteur video на французском языке.</span><span class="sxs-lookup"><span data-stu-id="44afe-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vidéo".</span></span>

> [!NOTE]
> <span data-ttu-id="44afe-157">Вы можете указать только одно переопределение на язык, в том числе для языкового стандарта по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="44afe-157">You may only specify a single override per language, including for the default locale.</span></span> <span data-ttu-id="44afe-158">Например, если по умолчанию используется языковой стандарт `en-us`, невозможно также указать переопределение для `en-us`.</span><span class="sxs-lookup"><span data-stu-id="44afe-158">For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="44afe-p113">В приведенном ниже примере применяется переопределение языкового стандарта для элемента [Description]. Сначала он задает языковой стандарт по умолчанию `en-us` и описание на английском языке, а затем указывает оператор [Override] с описанием на французском языке для языкового стандарта `fr-fr`:</span><span class="sxs-lookup"><span data-stu-id="44afe-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook."/>
</Description>
```

<span data-ttu-id="44afe-p114">Это значит, что надстройка предполагает языковой стандарт `en-us` по умолчанию. Пользователи увидят описание на английском языке в атрибуте `DefaultValue` для всех языковых стандартов, если на клиентском компьютере не выбран языковой стандарт `fr-fr`. В этом случае они увидят описание на французском языке.</span><span class="sxs-lookup"><span data-stu-id="44afe-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="44afe-p115">В следующем примере надстройка задает отдельное приложение, которое больше подходит для языкового стандарта и региональных параметров `fr-fr`. Пользователи видят изображение DefaultLogo.png по умолчанию, кроме тех случаев, когда на клиентском компьютере используется языковой стандарт `fr-fr`. В этом случае пользователи видят изображение FrenchLogo.png.</span><span class="sxs-lookup"><span data-stu-id="44afe-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="44afe-p116">В примере ниже показано, как локализовать ресурс в разделе `Resources`. Здесь применяется переопределение локали для изображения, и используется изображение, более подходящее для языка и региональных параметров `ja-jp`.</span><span class="sxs-lookup"><span data-stu-id="44afe-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="44afe-p117">В случае элемента [SourceLocation] поддержка дополнительных языковых стандартов означает предоставление отдельного исходного HTML-файла для каждого из указанных языковых стандартов. Пользователи с заданными языковыми стандартами увидят настраиваемые для них веб-страницы.</span><span class="sxs-lookup"><span data-stu-id="44afe-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="44afe-p118">В случае надстроек Outlook элемент [SourceLocation] также сопоставляется с форм-фактором. Это позволяет предоставлять отдельный локализованный исходный HTML-файл для каждого соответствующего форм-фактора. Вы можете задать один или несколько дочерних элементов [Override] в каждом применимом элементе параметров ([DesktopSettings], [TabletSettings] или [PhoneSettings]). В приведенном ниже примере показаны элементы параметров для форм-факторов настольного компьютера, планшета и смартфона, каждому из которых соответствует один HTML-файл для языкового стандарта по умолчанию и другой файл для французского языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="44afe-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


```xml
<DesktopSettings>
   <SourceLocation DefaultValue="https://contoso.com/Desktop.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Desktop.html" />
   </SourceLocation>
   <RequestedHeight>250</RequestedHeight>
</DesktopSettings>
<TabletSettings>
   <SourceLocation DefaultValue="https://contoso.com/Tablet.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Tablet.html" />
   </SourceLocation>
   <RequestedHeight>200</RequestedHeight>
</TabletSettings>
<PhoneSettings>
   <SourceLocation DefaultValue="https://contoso.com/Mobile.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Mobile.html" />
   </SourceLocation>
</PhoneSettings>
```

## <a name="localize-extended-overrides"></a><span data-ttu-id="44afe-174">Локализовать расширенные переопределения</span><span class="sxs-lookup"><span data-stu-id="44afe-174">Localize extended overrides</span></span>

<span data-ttu-id="44afe-175">Некоторые функции Office надстройки, например ярлыки клавиатуры, настраиваются с помощью файлов JSON, которые находятся на сервере, а не с XML-манифестом надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-175">Some extensibility features of Office Add-ins, such as keyboard shortcuts, are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span> <span data-ttu-id="44afe-176">В этом разделе предполагается, что вы знакомы с расширенными переопределениями.</span><span class="sxs-lookup"><span data-stu-id="44afe-176">This section assumes that you're familiar with extended overrides.</span></span> <span data-ttu-id="44afe-177">См. в этой ссылке Работа с расширенными [переопределениями элемента манифеста](extended-overrides.md) и [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="44afe-177">See [Work with extended overrides of the manifest](extended-overrides.md) and [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span>

<span data-ttu-id="44afe-178">Используйте атрибут `ResourceUrl` элемента [ExtendedOverrides,](../reference/manifest/extendedoverrides.md) чтобы указать Office файлу локализованных ресурсов.</span><span class="sxs-lookup"><span data-stu-id="44afe-178">Use the `ResourceUrl` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element to point Office to a file of localized resources.</span></span> <span data-ttu-id="44afe-179">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="44afe-179">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="44afe-180">Расширенный переопределяемый файл использует маркеры вместо строк.</span><span class="sxs-lookup"><span data-stu-id="44afe-180">The extended overrides file then uses tokens instead of strings.</span></span> <span data-ttu-id="44afe-181">Строки имен маркеров в файле ресурса.</span><span class="sxs-lookup"><span data-stu-id="44afe-181">The tokens name strings in the resource file.</span></span> <span data-ttu-id="44afe-182">Ниже приводится пример, который назначает клавишу ярлыка функции (определенной в другом месте), отображаемой области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-182">The following is an example that assigns a keyboard shortcut to a function (defined elsewhere) that displays the add-in's task pane.</span></span> <span data-ttu-id="44afe-183">Обратите внимание на эту разметку:</span><span class="sxs-lookup"><span data-stu-id="44afe-183">Note about this markup:</span></span>

- <span data-ttu-id="44afe-184">Пример не совсем допустимый.</span><span class="sxs-lookup"><span data-stu-id="44afe-184">The example isn't quite valid.</span></span> <span data-ttu-id="44afe-185">(Мы добавляем необходимое дополнительное свойство к ней ниже.)</span><span class="sxs-lookup"><span data-stu-id="44afe-185">(We add a required additional property to it below.)</span></span>
- <span data-ttu-id="44afe-186">Маркеры должны иметь формат **${resource.*name-of-resource*}**.</span><span class="sxs-lookup"><span data-stu-id="44afe-186">The tokens must have the format **${resource.*name-of-resource*}**.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ] 
}
```

<span data-ttu-id="44afe-187">Файл ресурса, который также форматирован JSON, имеет свойство верхнего уровня, которое делится на под свойства `resources` по локалу.</span><span class="sxs-lookup"><span data-stu-id="44afe-187">The resource file, which is also JSON-formatted, has a top-level `resources` property that is divided into subproperties by locale.</span></span> <span data-ttu-id="44afe-188">Для каждого локального адреса строка назначена каждому маркеру, который использовался в расширенном переопределяемом файле.</span><span class="sxs-lookup"><span data-stu-id="44afe-188">For each locale, a string is assigned to each token that was used in the extended overrides file.</span></span> <span data-ttu-id="44afe-189">Ниже приводится пример, в котором есть строки `en-us` для и `fr-fr` .</span><span class="sxs-lookup"><span data-stu-id="44afe-189">The following is an example which has strings for `en-us` and `fr-fr`.</span></span> <span data-ttu-id="44afe-190">В этом примере ярлык клавиатуры одинаковый в обоих локальных местах, но это не всегда так, особенно при локализации для локализованных локалов, которые имеют другой алфавит или систему записи, а значит, и другую клавиатуру.</span><span class="sxs-lookup"><span data-stu-id="44afe-190">In this example, the keyboard shortcut is the same in both locales, but that won't always be the case, especially when you are localizing for locales that have a different alphabet or writing system, and hence a different keyboard.</span></span>

```json
{
    "resources":{ 
        "en-us": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            }, 
        },
        "fr-fr": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Afficher le volet de tâche pour add-in",
              } 
        }
    }
}
```

<span data-ttu-id="44afe-191">В файле нет свойства одноранговой и `default` `en-us` `fr-fr` разделов.</span><span class="sxs-lookup"><span data-stu-id="44afe-191">There is no `default` property in the file that is a peer to the `en-us` and `fr-fr` sections.</span></span> <span data-ttu-id="44afe-192">Это происходит потому, что строки по умолчанию, которые используются, когда локализ Office хост-приложения не совпадает ни с одним из свойств *ll-cc* в файле *ресурсов,* должны быть определены в самом расширенном переопределяемом файле .</span><span class="sxs-lookup"><span data-stu-id="44afe-192">This is because the default strings, which are used when the locale of the Office host application doesn't match any of the *ll-cc* properties in the resources file, *must be defined in the extended overrides file itself*.</span></span> <span data-ttu-id="44afe-193">Определение строк по умолчанию непосредственно в расширенном переопределяемом файле гарантирует, что Office не скачивает файл ресурса, если локал приложения Office соответствует локальному стандарту надстройки (как указано в манифесте).</span><span class="sxs-lookup"><span data-stu-id="44afe-193">Defining the default strings directly in the extended overrides file ensures that Office doesn't download the resource file when the locale of the Office application matches the default locale of the add-in (as specified in the manifest).</span></span> <span data-ttu-id="44afe-194">Ниже приводится исправленная версия предыдущего примера расширенного переопределяемого файла, использующего маркеры ресурсов.</span><span class="sxs-lookup"><span data-stu-id="44afe-194">The following is a corrected version of the preceding example of an extended overrides file that uses resource tokens.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ],
    "resources": { 
        "default": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            } 
        }
    }
}
```

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="44afe-195">Приведение формата даты и времени к языковым параметрам клиента</span><span class="sxs-lookup"><span data-stu-id="44afe-195">Match date/time format with client locale</span></span>

<span data-ttu-id="44afe-196">Вы можете получить локализовку пользовательского интерфейса клиентского приложения Office с помощью **[свойства displayLanguage.]**</span><span class="sxs-lookup"><span data-stu-id="44afe-196">You can get the locale of the user interface of the Office client application by using the **[displayLanguage]** property.</span></span> <span data-ttu-id="44afe-197">Затем можно отобразить значения даты и времени в формате, соответствующем текущему Office приложения.</span><span class="sxs-lookup"><span data-stu-id="44afe-197">You can then display date and time values in a format consistent with the current locale of the Office application.</span></span> <span data-ttu-id="44afe-198">Один из способов сделать это — подготовить файл ресурсов, в котором задан формат отображения даты и времени для использования с каждым из языковых параметров, поддерживаемых Надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="44afe-198">One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports.</span></span> <span data-ttu-id="44afe-199">Во время запуска надстройка может использовать файл ресурсов и соответствовать соответствующему формату даты и времени с локализом, полученным из **[свойства displayLanguage.]**</span><span class="sxs-lookup"><span data-stu-id="44afe-199">At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the **[displayLanguage]** property.</span></span>

<span data-ttu-id="44afe-200">Вы можете получить локалику данных клиентского приложения Office с помощью [свойства contentLanguage.]</span><span class="sxs-lookup"><span data-stu-id="44afe-200">You can get the locale of the data of the Office client application by using the [contentLanguage] property.</span></span> <span data-ttu-id="44afe-201">На основе этого значения можно интерпретировать или отображать строки даты и времени.</span><span class="sxs-lookup"><span data-stu-id="44afe-201">Based on this value, you can then appropriately interpret or display date/time strings.</span></span> <span data-ttu-id="44afe-202">Например, в языковом стандарте `jp-JP` дата и время выражаются так: `yyyy/MM/dd`, а в языковом стандарте `fr-FR` так: `dd/MM/yyyy`.</span><span class="sxs-lookup"><span data-stu-id="44afe-202">For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="44afe-203">Использование Ajax для глобализации и локализации</span><span class="sxs-lookup"><span data-stu-id="44afe-203">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="44afe-204">Если для создания Надстройки Office вы используете Visual Studio, платформа .NET Framework и Ajax предоставляют способы глобализации и локализации файлов клиентских скриптов.</span><span class="sxs-lookup"><span data-stu-id="44afe-204">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="44afe-p127">Можно глобализировать и использовать расширения типов JavaScript [Date](/previous-versions/bb310850(v=vs.140)) и [Number](/previous-versions/bb310835(v=vs.140)) и объект JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) в коде JavaScript для Надстройка Office, чтобы отображать значения в зависимости от языковых параметров, заданных в текущем браузере. Дополнительные сведения см. в статье [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140)).</span><span class="sxs-lookup"><span data-stu-id="44afe-p127">You can globalize and use the [Date](/previous-versions/bb310850(v=vs.140)) and [Number](/previous-versions/bb310835(v=vs.140)) JavaScript type extensions and the JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140)).</span></span>

<span data-ttu-id="44afe-p128">Можно включить локализованные строки ресурсов напрямую в отдельные файлы JavaScript, чтобы предоставить клиентские файлы скриптов для разных языковых параметров, задаваемых в браузере или предоставляемых пользователем. Создайте отдельный файл скрипта для каждого поддерживаемого языкового параметра. В каждый файл скрипта включите объект в формате JSON, содержащий строки ресурсов для соответствующего языкового параметра. Локализованные значения применяются во время выполнения скрипта в браузере.</span><span class="sxs-lookup"><span data-stu-id="44afe-p128">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span>


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="44afe-211">Пример. Создание локализованной надстройки Office</span><span class="sxs-lookup"><span data-stu-id="44afe-211">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="44afe-212">В этом разделе представлены примеры того, как локализовать описание, отображаемое имя и пользовательский интерфейс Надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="44afe-212">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span> 

> [!NOTE]
> <span data-ttu-id="44afe-213">Чтобы скачать Visual Studio 2019 г., см. страницу [Visual Studio IDE.](https://visualstudio.microsoft.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="44afe-213">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="44afe-214">Во время установки потребуется выбрать рабочую нагрузку разработки для Office и SharePoint.</span><span class="sxs-lookup"><span data-stu-id="44afe-214">During installation you'll need to select the Office/SharePoint development workload.</span></span>

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="44afe-215">Настройка Office на использование дополнительных языков для отображения или редактирования</span><span class="sxs-lookup"><span data-stu-id="44afe-215">Configure Office to use additional languages for display or editing</span></span>

<span data-ttu-id="44afe-216">Чтобы запустить предоставленный пример кода, Office на компьютере, чтобы использовать дополнительные языки, чтобы можно было протестировать надстройку, переключив язык, используемый для отображения в меню и командах, для редактирования и проверки или обоих.</span><span class="sxs-lookup"><span data-stu-id="44afe-216">To run the sample code provided, configure Office on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="44afe-217">Для установки дополнительного языка можно использовать языковой пакет Office.</span><span class="sxs-lookup"><span data-stu-id="44afe-217">You can use an Office Language pack to install an additional language.</span></span> <span data-ttu-id="44afe-218">Дополнительные сведения о языковых пакетах и способах их получения см. на странице [дополнительных языковых пакетов для Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span><span class="sxs-lookup"><span data-stu-id="44afe-218">For more information about Language Packs and where to get them, see [Language Accessory Pack for Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span></span>

<span data-ttu-id="44afe-219">После установки языкового пакета вы можете настроить Office на использование установленного языка для пользовательского интерфейса и/или для редактирования содержимого документов.</span><span class="sxs-lookup"><span data-stu-id="44afe-219">After you install the Language Accessory Pack, you can configure Office to use the installed language for display in the UI, for editing document content, or both.</span></span> <span data-ttu-id="44afe-220">В примере в этой статье используется установка Office, в которой применяется испанский языковой пакет.</span><span class="sxs-lookup"><span data-stu-id="44afe-220">The example in this article uses an installation of Office that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="44afe-221">Создание проекта надстройки Office</span><span class="sxs-lookup"><span data-stu-id="44afe-221">Create an Office Add-in project</span></span>

<span data-ttu-id="44afe-222">Необходимо создать проект Visual Studio 2019 Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-222">You'll need to create a Visual Studio 2019 Office Add-in project.</span></span>

> [!NOTE]
> <span data-ttu-id="44afe-223">Если вы не установили Visual Studio 2019 г., см. на странице [Visual Studio IDE](https://visualstudio.microsoft.com/vs/) для инструкций по загрузке.</span><span class="sxs-lookup"><span data-stu-id="44afe-223">If you haven't installed Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/) for download instructions.</span></span> <span data-ttu-id="44afe-224">Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint.</span><span class="sxs-lookup"><span data-stu-id="44afe-224">During installation you'll need to select the Office/SharePoint development workload.</span></span> <span data-ttu-id="44afe-225">Если вы установили Visual Studio 2019 [г.,](/visualstudio/install/modify-visual-studio/) используйте Visual Studio Installer, чтобы обеспечить Office/SharePoint разработки.</span><span class="sxs-lookup"><span data-stu-id="44afe-225">If you have previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio/) to ensure that the Office/SharePoint development workload is installed.</span></span>

1. <span data-ttu-id="44afe-226">Выберите **Создание нового проекта**.</span><span class="sxs-lookup"><span data-stu-id="44afe-226">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="44afe-227">Используя поле поиска, введите **надстройка**.</span><span class="sxs-lookup"><span data-stu-id="44afe-227">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="44afe-228">Выберите вариант **Веб-надстройка Word** и нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="44afe-228">Choose **Word Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="44afe-229">Назови свой **проект WorldReadyAddIn и** выберите **Create**.</span><span class="sxs-lookup"><span data-stu-id="44afe-229">Name your project **WorldReadyAddIn** and select **Create**.</span></span>

4. <span data-ttu-id="44afe-p134">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="44afe-p134">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>


### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="44afe-232">Локализация текста, используемого в вашей надстройке</span><span class="sxs-lookup"><span data-stu-id="44afe-232">Localize the text used in your add-in</span></span>

<span data-ttu-id="44afe-233">Текст, который нужно локализовать на другом языке, отображается в двух областях:</span><span class="sxs-lookup"><span data-stu-id="44afe-233">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="44afe-p135">**Отображаемое имя и описание надстройки**. Они управляются записями в файле манифеста приложения.</span><span class="sxs-lookup"><span data-stu-id="44afe-p135">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>

-  <span data-ttu-id="44afe-236">**Пользовательский интерфейс надстройки**.</span><span class="sxs-lookup"><span data-stu-id="44afe-236">**Add-in UI**.</span></span> <span data-ttu-id="44afe-237">Вы можете локализовать строки, отображаемые в пользовательском интерфейсе надстройки, с помощью кода JavaScript, например используя отдельный файл ресурсов с локализованными строками.</span><span class="sxs-lookup"><span data-stu-id="44afe-237">You can localize the strings that appear in your add-in UI by using JavaScript code, for example, by using a separate resource file that contains the localized strings.</span></span>

<span data-ttu-id="44afe-238">Локализация отображаемого имени и описания надстройки:</span><span class="sxs-lookup"><span data-stu-id="44afe-238">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="44afe-239">В **обозревателе решений** разверните узлы **WorldReadyAddIn** и **WorldReadyAddInManifest**, а затем выберите **WorldReadyAddIn.xml**.</span><span class="sxs-lookup"><span data-stu-id="44afe-239">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose **WorldReadyAddIn.xml**.</span></span>

2. <span data-ttu-id="44afe-240">В WorldReadyAddInManifest.xml замените [элементы DisplayName] и [Description] следующим блоком кода.</span><span class="sxs-lookup"><span data-stu-id="44afe-240">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code.</span></span>

    > [!NOTE]
    > <span data-ttu-id="44afe-241">Вы можете заменить локализованные строки на испанском языке, используемые в этом примере для элементов [DisplayName] и [Description], локализованными строками на любом другом языке.</span><span class="sxs-lookup"><span data-stu-id="44afe-241">You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="44afe-242">После изменения отображаемого языка для Office 2013, к примеру, с английского на испанский и последующего запуска надстройки отображаемое имя и описание надстройки локализуются.</span><span class="sxs-lookup"><span data-stu-id="44afe-242">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span>

<span data-ttu-id="44afe-243">Настройка пользовательского интерфейса надстройки:</span><span class="sxs-lookup"><span data-stu-id="44afe-243">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="44afe-244">В **обозревателе решений** Visual Studio выберите элемент **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="44afe-244">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>

2. <span data-ttu-id="44afe-245">Замените содержимое элемента `<body>` в файле Home.html на приведенный ниже HTML-код и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="44afe-245">Replace the `<body>` element contents in Home.html with the following HTML, and save the file.</span></span>

    ```html
    <body>
        <!-- Page content -->
        <div id="content-header" class="ms-bgColor-themePrimary ms-font-xl">
            <div class="padding">
                <h1 id="greeting" class="ms-fontColor-white"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div class="ms-font-m">
                    <p id="about"></p>
                </div>
            </div>
        </div>
    </body>
    ```

<span data-ttu-id="44afe-246">На приведенном ниже рисунке показаны элемент заголовка (h1) и элемент абзаца (p), в которых будет отображаться локализованный текст после завершения оставшихся действий и запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-246">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when you complete the remaining steps and run the add-in.</span></span>

<span data-ttu-id="44afe-247">*Рис. 1. Пользовательский интерфейс надстройки*</span><span class="sxs-lookup"><span data-stu-id="44afe-247">*Figure 1. The add-in UI*</span></span>

![Пользовательский интерфейс приложения с выделенными разделами.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="44afe-249">Добавление файла ресурсов с локализованными строками</span><span class="sxs-lookup"><span data-stu-id="44afe-249">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="44afe-250">Файл ресурсов JavaScript содержит строки, используемые для пользовательского интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-250">The JavaScript resource file contains the strings used for the add-in UI.</span></span> <span data-ttu-id="44afe-251">HTML-код для пользовательского интерфейса примера надстройки содержит элемент `<h1>`, отображающий приветствие, и элемент `<p>`, который знакомит пользователя с надстройкой.</span><span class="sxs-lookup"><span data-stu-id="44afe-251">The HTML for the sample add-in UI contains an `<h1>` element that displays a greeting, and a `<p>` element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="44afe-p138">Чтобы включить локализованные строки для заголовка и абзаца, нужно поместить строки в отдельный файл ресурса. Файл ресурса создает объект JavaScript, который содержит отдельный объект Нотация объектов JavaScript (JSON) для каждого набора локализованных строк. Файл ресурса также предоставляет метод для получения соответствующего объекта JSON для определенного региона.</span><span class="sxs-lookup"><span data-stu-id="44afe-p138">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span>

<span data-ttu-id="44afe-255">Добавление файла ресурсов в проект надстройки:</span><span class="sxs-lookup"><span data-stu-id="44afe-255">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="44afe-256">В **обозревателе решений** Visual Studio, щелкните правой кнопкой мыши проект **WorldReadyAddInWeb** и выберите **Добавить** > **Создать элемент**.</span><span class="sxs-lookup"><span data-stu-id="44afe-256">In **Solution Explorer** in Visual Studio, right-click the **WorldReadyAddInWeb** project and choose **Add** > **New Item**.</span></span> 

2. <span data-ttu-id="44afe-257">В диалоговом окне **Добавление нового элемента** выберите параметр **файл JavaScript**.</span><span class="sxs-lookup"><span data-stu-id="44afe-257">In the **Add New Item** dialog box, choose **JavaScript File**.</span></span>

3. <span data-ttu-id="44afe-258">Введите **UIStrings.js** в качестве имени файла и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="44afe-258">Enter **UIStrings.js** as the file name and choose **Add**.</span></span>

4. <span data-ttu-id="44afe-259">Добавьте в файл UIStrings.js приведенный ниже код и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="44afe-259">Add the following code to the UIStrings.js file, and save the file.</span></span>

    ```js
    /* Store the locale-specific strings */

    var UIStrings = (function ()
    {
        "use strict";

        var UIStrings = {};

        // JSON object for English strings
        UIStrings.EN =
        {
            "Greeting": "Welcome",
            "Introduction": "This is my localized add-in."
        };

        // JSON object for Spanish strings
        UIStrings.ES =
        {
            "Greeting": "Bienvenido",
            "Introduction": "Esta es mi aplicación localizada."
        };

        UIStrings.getLocaleStrings = function (locale)
        {
            var text;

            // Get the resource strings that match the language.
            switch (locale)
            {
                case 'en-US':
                    text = UIStrings.EN;
                    break;
                case 'es-ES':
                    text = UIStrings.ES;
                    break;
                default:
                    text = UIStrings.EN;
                    break;
            }

            return text;
        };

        return UIStrings;
    })();
    ```

<span data-ttu-id="44afe-260">Файл ресурса UIStrings.js создает объект **UIStrings**, который содержит локализованные строки пользовательского интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-260">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span>

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="44afe-261">Локализация текста, используемого для пользовательского интерфейса надстройки</span><span class="sxs-lookup"><span data-stu-id="44afe-261">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="44afe-p139">Чтобы использовать в надстройке файл ресурсов, вам потребуется добавить для него тег сценария в файл Home.html. При загрузке файла Home.html выполняется файл UIStrings.js, и объект **UIStrings**, используемый для получения строк, становится доступен в коде. Добавьте приведенный ниже HTML-код в тег заголовка для файла Home.html, чтобы сделать объект **UIStrings** доступным в коде.</span><span class="sxs-lookup"><span data-stu-id="44afe-p139">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="44afe-265">Теперь вы можете использовать объект **UIStrings**, чтобы задать строки для пользовательского интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="44afe-265">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="44afe-266">Если вы хотите изменить локализацию надстройки в зависимости от языка, используемого для отображения в меню и командах в клиентском приложении Office, вы используете **свойство Office.context.displayLanguage** для получения языка для этого языка.</span><span class="sxs-lookup"><span data-stu-id="44afe-266">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the Office client application, you use the **Office.context.displayLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="44afe-267">Например, если язык приложения использует испанский для отображения в меню и командах, **свойство Office.context.displayLanguage** возвращает языковой код es-ES.</span><span class="sxs-lookup"><span data-stu-id="44afe-267">For example, if the application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="44afe-268">Если вы хотите изменить локализацию надстройки в зависимости от языка, используемого для редактирования контента документов, вы используете **свойство Office.context.contentLanguage** для получения языка для этого языка.</span><span class="sxs-lookup"><span data-stu-id="44afe-268">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the **Office.context.contentLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="44afe-269">Например, если язык приложений использует испанский для редактирования контента документов, **свойство Office.context.contentLanguage** возвращает языковой код es-ES.</span><span class="sxs-lookup"><span data-stu-id="44afe-269">For example, if the application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="44afe-270">После получения языка, используемого приложением, вы можете использовать **UIStrings** для получения набора локализованных строк, которые совпадают с языком приложений.</span><span class="sxs-lookup"><span data-stu-id="44afe-270">After you know the language the application is using, you can use **UIStrings** to get the set of localized strings that matches the application language.</span></span>

<span data-ttu-id="44afe-271">Замените код в файле Home.js на следующий код.</span><span class="sxs-lookup"><span data-stu-id="44afe-271">Replace the code in the Home.js file with the following code.</span></span> <span data-ttu-id="44afe-272">В коде показано, как можно изменять строки, используемые в элементах пользовательского интерфейса Home.html на основе языка отображения приложения или языка редактирования приложения.</span><span class="sxs-lookup"><span data-stu-id="44afe-272">The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the application or the editing language of the application.</span></span>

> [!NOTE]
> <span data-ttu-id="44afe-273">Чтобы переключаться между локализацией надстройки, основанной на языке редактирования, удалите символы комментария из строки кода `var myLanguage = Office.context.contentLanguage;` и заключите в знаки комментария строку кода `var myLanguage = Office.context.displayLanguage;`.</span><span class="sxs-lookup"><span data-stu-id="44afe-273">To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the Office application.
            var myLanguage = Office.context.displayLanguage;
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Introduction);
        });
    };
})();
```

### <a name="test-your-localized-add-in"></a><span data-ttu-id="44afe-274">Тестирование локализованной надстройки</span><span class="sxs-lookup"><span data-stu-id="44afe-274">Test your localized add-in</span></span>

<span data-ttu-id="44afe-275">Чтобы проверить локализованную надстройку, измените язык, используемый для отображения или редактирования в приложении Office и запустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="44afe-275">To test your localized add-in, change the language used for display or editing in the Office application and then run your add-in.</span></span>

<span data-ttu-id="44afe-276">Изменение языка, используемого для отображения или редактирования в надстройке:</span><span class="sxs-lookup"><span data-stu-id="44afe-276">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="44afe-277">В Word выберите **Файл** > **Параметры** > **Язык**.</span><span class="sxs-lookup"><span data-stu-id="44afe-277">In Word, choose **File** > **Options** > **Language**.</span></span> <span data-ttu-id="44afe-278">На рисунке ниже показано диалоговое окно **Параметры Word**, открытое на вкладке языка.</span><span class="sxs-lookup"><span data-stu-id="44afe-278">The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>

    <span data-ttu-id="44afe-279">*Рис. 2. Параметры языка в диалоговом окне "Параметры Word"*</span><span class="sxs-lookup"><span data-stu-id="44afe-279">*Figure 2. Language options in the Word Options dialog box*</span></span>

    ![Диалоговое окно Word Options.](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="44afe-281">В разделе **Выбор языков интерфейса** выберите язык, на котором должны отображаться данные (например, испанский), а затем нажмите стрелку вверх, чтобы переместить испанский язык в начало списка.</span><span class="sxs-lookup"><span data-stu-id="44afe-281">Under **Choose Display Language**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list.</span></span> <span data-ttu-id="44afe-282">Кроме того, чтобы изменить язык, используемый для редактирования, в статье **Выберите** языки редактирования выберите язык, который необходимо использовать для редактирования, например испанский, а затем выберите **Set as Default**.</span><span class="sxs-lookup"><span data-stu-id="44afe-282">Alternatively, to change the language used for editing, under **Choose Editing Languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>

3. <span data-ttu-id="44afe-283">Нажмите кнопку **ОК**, чтобы подтвердить выбор, а затем закройте Word.</span><span class="sxs-lookup"><span data-stu-id="44afe-283">Choose **OK** to confirm your selection, and then close Word.</span></span>

4. <span data-ttu-id="44afe-284">Нажмите клавишу **F5** в Visual Studio, чтобы запустить пример надстройки, или выберите **Отладка** > **Начать отладку** в строке меню.</span><span class="sxs-lookup"><span data-stu-id="44afe-284">Press **F5** in Visual Studio to run the sample add-in, or choose **Debug** > **Start Debugging** from the menu bar.</span></span>

5. <span data-ttu-id="44afe-285">В Word выберите **Главная** > **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="44afe-285">In Word, choose **Home** > **Show Taskpane**.</span></span>

<span data-ttu-id="44afe-286">После запуска строки в пользовательском интерфейсе надстройки изменяются в соответствии с языком, используемым приложением, как показано на следующем рисунке.</span><span class="sxs-lookup"><span data-stu-id="44afe-286">Once running, the strings in the add-in UI change to match the language used by the application, as shown in the following figure.</span></span>


<span data-ttu-id="44afe-287">*Рис. 3. Пользовательский интерфейс надстройки с локализованным текстом*</span><span class="sxs-lookup"><span data-stu-id="44afe-287">*Figure 3. Add-in UI with localized text*</span></span>

![Приложение с локализованным текстом пользовательского интерфейса.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="44afe-289">См. также</span><span class="sxs-lookup"><span data-stu-id="44afe-289">See also</span></span>

- [<span data-ttu-id="44afe-290">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="44afe-290">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- <span data-ttu-id="44afe-291">[Идентификаторы языков и значения OptionState Id в Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="44afe-291">[Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span></span>

[DefaultLocale]:         ../reference/manifest/defaultlocale.md
[Описание]:           ../reference/manifest/description.md
[Description]:           ../reference/manifest/description.md
[DisplayName]:           ../reference/manifest/displayname.md
[IconUrl]:               ../reference/manifest/iconurl.md
[HighResolutionIconUrl]: ../reference/manifest/highresolutioniconurl.md
[Resources]:             ../reference/manifest/resources.md
[SourceLocation]:        ../reference/manifest/sourcelocation.md
[Override]:              ../reference/manifest/override.md
[DesktopSettings]:       ../reference/manifest/desktopsettings.md
[TabletSettings]:        ../reference/manifest/tabletsettings.md
[PhoneSettings]:         ../reference/manifest/phonesettings.md
[displayLanguage]:       /javascript/api/office/office.context#displaylanguage
[contentLanguage]:       /javascript/api/office/office.context#contentlanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
