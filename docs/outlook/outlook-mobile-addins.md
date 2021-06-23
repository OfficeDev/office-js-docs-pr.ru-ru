---
title: Надстройки Outlook для Outlook Mobile
description: Outlook мобильные надстройки поддерживаются во всех Microsoft 365 учетных записях, Outlook.com, и поддержка скоро появится на учетных записях gmail.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: e4f44d36f6dab5841d249e3754e78c5dd7a3813e
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076737"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="6c8ba-103">Надстройки для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="6c8ba-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="6c8ba-p101">В Outlook Mobile теперь работают надстройки, использующие те же API, что и в других конечных точках Outlook. Если вы уже создали надстройку для Outlook, вам будет легко запустить ее в Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="6c8ba-106">Outlook мобильные надстройки поддерживаются во всех учетных записях Microsoft 365, Outlook.com, а поддержка скоро появится на учетных записях Gmail.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-106">Outlook mobile add-ins are supported on all Microsoft 365 business accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="6c8ba-107">**Пример области задач в Outlook для iOS**</span><span class="sxs-lookup"><span data-stu-id="6c8ba-107">**An example task pane in Outlook on iOS**</span></span>

![Снимок экрана области задач в Outlook iOS.](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="6c8ba-109">**Пример области задач в Outlook для Android**</span><span class="sxs-lookup"><span data-stu-id="6c8ba-109">**An example task pane in Outlook on Android**</span></span>

![Снимок экрана области задач в Outlook Android.](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="6c8ba-111">Надстройки не работают в современной версии Outlook в мобильном браузере.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="6c8ba-112">Дополнительные сведения см. [в Outlook в мобильном браузере.](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)</span><span class="sxs-lookup"><span data-stu-id="6c8ba-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="6c8ba-113">Чем отличаются надстройки для мобильных устройств?</span><span class="sxs-lookup"><span data-stu-id="6c8ba-113">What's different on mobile?</span></span>

- <span data-ttu-id="6c8ba-p103">Небольшой размер и скорость взаимодействия усложняют разработку для мобильных устройств. Чтобы пользователи получали только качественные приложения, мы устанавливаем строгие требования, которым должна соответствовать надстройка с заявленной поддержкой мобильных устройств для утверждения в AppSource.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p103">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
  - <span data-ttu-id="6c8ba-116">В надстройке **ДОЛЖНЫ** соблюдаться [рекомендации по пользовательскому интерфейсу](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="6c8ba-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
  - <span data-ttu-id="6c8ba-117">Сценарий для использования надстройки **ДОЛЖЕН** [быть уместным на мобильных устройствах](#what-makes-a-good-scenario-for-mobile-add-ins).</span><span class="sxs-lookup"><span data-stu-id="6c8ba-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="6c8ba-118">В общем, в настоящее время поддерживается только режим чтения сообщений.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="6c8ba-119">Это `MobileMessageReadCommandSurface` означает, что это единственное [приложение ExtensionPoint,](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) которое необходимо объявить в мобильном разделе манифеста.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="6c8ba-120">Однако режим организатора встречи поддерживается для встроенных надстроек поставщика собраний в Интернете, которые вместо этого объявляют точку расширения [MobileOnlineMeetingCommandSurface.](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface)</span><span class="sxs-lookup"><span data-stu-id="6c8ba-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface).</span></span> <span data-ttu-id="6c8ba-121">Дополнительные [статьи об этом сценарии см. в статье Создание](online-meeting.md) надстройки Outlook для поставщика веб-собраний.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="6c8ba-p105">API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) не поддерживается на мобильных устройствах, так как мобильное приложение использует интерфейсы REST API для связи с сервером. Если внутреннему серверу приложения требуется подключиться к серверу Exchange, вы можете совершать вызовы REST API с помощью маркера обратного вызова. Дополнительные сведения см. в статье [Использование интерфейсов REST API Outlook из надстройки Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p105">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="6c8ba-125">Отправляя надстройку в магазин с элементом [MobileFormFactor](../reference/manifest/mobileformfactor.md) в манифесте, необходимо принять условия приложения для разработчиков надстроек на iOS, а также указать свой идентификатор разработчика Apple для проверки.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="6c8ba-126">Кроме того, в манифесте необходимо объявить элемент `MobileFormFactor`, а также указать правильные [элементы управления](../reference/manifest/control.md) и [размеры значков](../reference/manifest/icon.md).</span><span class="sxs-lookup"><span data-stu-id="6c8ba-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="6c8ba-127">Для каких сценариев хорошо подходят мобильные надстройки?</span><span class="sxs-lookup"><span data-stu-id="6c8ba-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="6c8ba-p106">Помните, что средняя продолжительность сеанса Outlook на телефоне значительно ниже, чем на компьютере. Это означает, что надстройка должна работать быстро, позволяя пользователю зайти, выйти и вернуться к работе с электронной почтой.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p106">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="6c8ba-130">Ниже приведены примеры сценариев, для которых подходит Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="6c8ba-p107">Надстройка передает ценные сведения в Outlook, помогая пользователям сортировать свою почту и отвечать надлежащим образом. Пример: надстройка CRM, позволяющая пользователю просматривать сведения о клиентах и делиться соответствующей информацией.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p107">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="6c8ba-p108">Надстройка повышает ценность содержимого электронной почты пользователя, сохраняя сведения в системе отслеживания, совместной работы или другой подобной системе. Пример: надстройка, позволяющая пользователям преобразовывать электронные сообщения в элементы задач для отслеживания проектов или заявки в службу поддержки.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p108">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="6c8ba-135">**Пример действий пользователя для создания карточки Trello из электронного сообщения на iOS**</span><span class="sxs-lookup"><span data-stu-id="6c8ba-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![Анимированный GIF, показывающий взаимодействие пользователя с Outlook надстройки Mobile на iOS.](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="6c8ba-137">**Пример действий пользователя для создания карточки Trello из электронного сообщения на Android**</span><span class="sxs-lookup"><span data-stu-id="6c8ba-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Анимированный GIF, показывающий взаимодействие пользователя с Outlook надстройки Mobile на Android.](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="6c8ba-139">Тестирование надстроек на мобильных устройствах</span><span class="sxs-lookup"><span data-stu-id="6c8ba-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="6c8ba-p109">Чтобы протестировать надстройку в Outlook Mobile, вы можете загрузить неопубликованную надстройку в учетную запись Office 365 или Outlook.com. В Outlook в Интернете нажмите значок шестеренки и выберите пункт **Управление интеграцией** или **Управление надстройками**. В верхней части экрана нажмите надпись **Щелкните здесь, чтобы добавить пользовательскую надстройку** и отправьте манифест. Убедитесь, что манифест отформатирован надлежащим образом и включает `MobileFormFactor`. В противном случае он не загрузится.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p109">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="6c8ba-p110">Подготовив надстройку к работе, протестируйте ее на экранах различных размеров, в том числе на телефонах и планшетах. Убедитесь, что она соответствует требованиям к специальным возможностям: контрастности, размеру шрифта, а также возможности работы со средствами чтения с экрана, такими как VoiceOver в iOS и TalkBack в Android.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-p110">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="6c8ba-145">Устранение неполадок на мобильных устройствах может быть затруднено, так как у вас может не быть используемых средств.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-145">Troubleshooting on mobile can be hard since you may not have the tools you're used to.</span></span> <span data-ttu-id="6c8ba-146">Однако одним из вариантов устранения неполадок на iOS является использование Fiddler (ознакомьтесь с этим учебником по его использованию [с помощью устройства iOS).](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)</span><span class="sxs-lookup"><span data-stu-id="6c8ba-146">However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span></span>

## <a name="next-steps"></a><span data-ttu-id="6c8ba-147">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="6c8ba-147">Next steps</span></span>

<span data-ttu-id="6c8ba-148">Узнайте, как:</span><span class="sxs-lookup"><span data-stu-id="6c8ba-148">Learn how to:</span></span>

- <span data-ttu-id="6c8ba-149">[Добавить поддержку мобильных устройств в манифест надстройки](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="6c8ba-149">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="6c8ba-150">[Разработать отличный мобильный интерфейс для надстройки](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="6c8ba-150">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="6c8ba-151">[Получить маркер доступа и вызвать REST API Outlook](use-rest-api.md) из надстройки.</span><span class="sxs-lookup"><span data-stu-id="6c8ba-151">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
