---
title: Надстройки Outlook для Outlook Mobile
description: Надстройки Outlook Mobile поддерживаются во всех коммерческих учетных записях Office 365, Outlook.com. Скоро их можно будет использовать и в учетных записях Gmail.
ms.date: 04/13/2020
localization_priority: Normal
ms.openlocfilehash: 4b6341ac1b340ebc46c616ae4274bfdf1e2d0672
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241086"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="737ea-103">Надстройки для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="737ea-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="737ea-p101">В Outlook Mobile теперь работают надстройки, использующие те же API, что и в других конечных точках Outlook. Если вы уже создали надстройку для Outlook, вам будет легко запустить ее в Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="737ea-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="737ea-106">Надстройки Outlook Mobile поддерживаются во всех коммерческих учетных записях Office 365, Outlook.com. Скоро их можно будет использовать и в учетных записях Gmail.</span><span class="sxs-lookup"><span data-stu-id="737ea-106">Outlook mobile add-ins are supported on all Office 365 Commercial accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="737ea-107">**Пример области задач в Outlook для iOS**</span><span class="sxs-lookup"><span data-stu-id="737ea-107">**An example task pane in Outlook on iOS**</span></span>

![Снимок экрана с областью задач в Outlook для iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="737ea-109">**Пример области задач в Outlook для Android**</span><span class="sxs-lookup"><span data-stu-id="737ea-109">**An example task pane in Outlook on Android**</span></span>

![Снимок экрана с областью задач в Outlook для Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="737ea-111">Надстройки не работают в современной версии Outlook в браузере мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="737ea-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="737ea-112">Дополнительные сведения см. [в статье Outlook в браузере мобильного устройства](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span><span class="sxs-lookup"><span data-stu-id="737ea-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="737ea-113">Чем отличаются надстройки для мобильных устройств?</span><span class="sxs-lookup"><span data-stu-id="737ea-113">What's different on mobile?</span></span>

- <span data-ttu-id="737ea-p103">Небольшой размер и скорость взаимодействия усложняют разработку для мобильных устройств. Чтобы пользователи получали только качественные приложения, мы устанавливаем строгие требования, которым должна соответствовать надстройка с заявленной поддержкой мобильных устройств для утверждения в AppSource.</span><span class="sxs-lookup"><span data-stu-id="737ea-p103">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="737ea-116">В надстройке **ДОЛЖНЫ** соблюдаться [рекомендации по пользовательскому интерфейсу](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="737ea-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="737ea-117">Сценарий для использования надстройки **ДОЛЖЕН** [быть уместным на мобильных устройствах](#what-makes-a-good-scenario-for-mobile-add-ins).</span><span class="sxs-lookup"><span data-stu-id="737ea-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="737ea-118">Как правило, в настоящее время поддерживается только режим чтения сообщений.</span><span class="sxs-lookup"><span data-stu-id="737ea-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="737ea-119">Это означает `MobileMessageReadCommandSurface` единственный [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) , который следует объявить в разделе мобильного устройства манифеста.</span><span class="sxs-lookup"><span data-stu-id="737ea-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="737ea-120">Однако режим организатора встречи поддерживается для встроенных надстроек поставщика собраний по сети, которые вместо этого объявляют [точку расширения мобилеонлинемитингкоммандсурфаце](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span><span class="sxs-lookup"><span data-stu-id="737ea-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span></span> <span data-ttu-id="737ea-121">Для получения дополнительных сведений об этом сценарии обратитесь к статье [Создание надстройки Outlook для мобильных устройств](online-meeting.md) .</span><span class="sxs-lookup"><span data-stu-id="737ea-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="737ea-p105">API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) не поддерживается на мобильных устройствах, так как мобильное приложение использует интерфейсы REST API для связи с сервером. Если внутреннему серверу приложения требуется подключиться к серверу Exchange, вы можете совершать вызовы REST API с помощью маркера обратного вызова. Дополнительные сведения см. в статье [Использование интерфейсов REST API Outlook из надстройки Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="737ea-p105">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="737ea-125">Отправляя надстройку в магазин с элементом [MobileFormFactor](../reference/manifest/mobileformfactor.md) в манифесте, необходимо принять условия приложения для разработчиков надстроек на iOS, а также указать свой идентификатор разработчика Apple для проверки.</span><span class="sxs-lookup"><span data-stu-id="737ea-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="737ea-126">Кроме того, в манифесте необходимо объявить элемент `MobileFormFactor`, а также указать правильные [элементы управления](../reference/manifest/control.md) и [размеры значков](../reference/manifest/icon.md).</span><span class="sxs-lookup"><span data-stu-id="737ea-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="737ea-127">Для каких сценариев хорошо подходят мобильные надстройки?</span><span class="sxs-lookup"><span data-stu-id="737ea-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="737ea-p106">Помните, что средняя продолжительность сеанса Outlook на телефоне значительно ниже, чем на компьютере. Это означает, что надстройка должна работать быстро, позволяя пользователю зайти, выйти и вернуться к работе с электронной почтой.</span><span class="sxs-lookup"><span data-stu-id="737ea-p106">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="737ea-130">Ниже приведены примеры сценариев, для которых подходит Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="737ea-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="737ea-p107">Надстройка передает ценные сведения в Outlook, помогая пользователям сортировать свою почту и отвечать надлежащим образом. Пример: надстройка CRM, позволяющая пользователю просматривать сведения о клиентах и делиться соответствующей информацией.</span><span class="sxs-lookup"><span data-stu-id="737ea-p107">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="737ea-p108">Надстройка повышает ценность содержимого электронной почты пользователя, сохраняя сведения в системе отслеживания, совместной работы или другой подобной системе. Пример: надстройка, позволяющая пользователям преобразовывать электронные сообщения в элементы задач для отслеживания проектов или заявки в службу поддержки.</span><span class="sxs-lookup"><span data-stu-id="737ea-p108">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="737ea-135">**Пример действий пользователя для создания карточки Trello из электронного сообщения на iOS**</span><span class="sxs-lookup"><span data-stu-id="737ea-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![Анимированный GIF-файл, иллюстрирующий взаимодействие пользователя с надстройкой Outlook Mobile на iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="737ea-137">**Пример действий пользователя для создания карточки Trello из электронного сообщения на Android**</span><span class="sxs-lookup"><span data-stu-id="737ea-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Анимированный GIF-файл, иллюстрирующий взаимодействие пользователя с надстройкой Outlook Mobile на Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="737ea-139">Тестирование надстроек на мобильных устройствах</span><span class="sxs-lookup"><span data-stu-id="737ea-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="737ea-p109">Чтобы протестировать надстройку в Outlook Mobile, вы можете загрузить неопубликованную надстройку в учетную запись Office 365 или Outlook.com. В Outlook в Интернете нажмите значок шестеренки и выберите пункт **Управление интеграцией** или **Управление надстройками**. В верхней части экрана нажмите надпись **Щелкните здесь, чтобы добавить пользовательскую надстройку** и отправьте манифест. Убедитесь, что манифест отформатирован надлежащим образом и включает `MobileFormFactor`. В противном случае он не загрузится.</span><span class="sxs-lookup"><span data-stu-id="737ea-p109">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="737ea-p110">Подготовив надстройку к работе, протестируйте ее на экранах различных размеров, в том числе на телефонах и планшетах. Убедитесь, что она соответствует требованиям к специальным возможностям: контрастности, размеру шрифта, а также возможности работы со средствами чтения с экрана, такими как VoiceOver в iOS и TalkBack в Android.</span><span class="sxs-lookup"><span data-stu-id="737ea-p110">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="737ea-p111">Устранять неполадки на мобильных устройствах может быть сложно, так как в вашем распоряжении может не оказаться привычных инструментов. Один из вариантов устранения неполадок — [использование Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). А если вы уже использовали Fiddler, ознакомьтесь с [этим руководством по его использованию с устройствами iOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).</span><span class="sxs-lookup"><span data-stu-id="737ea-p111">Troubleshooting on mobile can be hard since you may not have the tools you're used to. One option for troubleshooting is to [use Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). Or, if you've used Fiddler before, check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).</span></span>

## <a name="next-steps"></a><span data-ttu-id="737ea-148">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="737ea-148">Next steps</span></span>

<span data-ttu-id="737ea-149">Узнайте, как:</span><span class="sxs-lookup"><span data-stu-id="737ea-149">Learn how to:</span></span>

- <span data-ttu-id="737ea-150">[Добавить поддержку мобильных устройств в манифест надстройки](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="737ea-150">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="737ea-151">[Разработать отличный мобильный интерфейс для надстройки](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="737ea-151">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="737ea-152">[Получить маркер доступа и вызвать REST API Outlook](use-rest-api.md) из надстройки.</span><span class="sxs-lookup"><span data-stu-id="737ea-152">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
