---
title: Наборы обязательных элементов API JavaScript для Outlook
description: Узнайте больше о наборах обязательных элементов API JavaScript для Outlook
ms.date: 05/17/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 49cfcfee075ba01f077162cef415ed58211b95f6
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936555"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Outlook

Для надстроек Outlook требуются определенные версии API, которые указываются в элементе Requirements в их манифесте. Надстройки Outlook всегда включают элемент Set с атрибутом , для которого задано значение , и атрибутом , для которого установлен минимальный набор обязательных элементов API, поддерживающий сценарии надстройки.

Например, в следующем фрагменте манифеста указан минимальный набор обязательных элементов 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Все API-интерфейсы Outlook приведены в [наборе обязательных элементов](../../develop/specify-office-hosts-and-api-requirements.md) `Mailbox`. У набора обязательных элементов`Mailbox` есть версии, а каждый новый выпускаемый набор API-интерфейсов приведен в наборе более поздней версии. Не все клиенты Outlook поддерживают новейший набор API-интерфейсов, но если для клиента Outlook объявлена поддержка набора обязательных элементов, обычно он поддерживает все API-интерфейсы в этом наборе (ознакомьтесь с документацией по конкретному API или функции на наличие исключений).

Задайте версию минимального набора обязательных элементов в манифесте, чтобы указать клиент Outlook, в котором появится надстройка. Если клиент не поддерживает минимальный набор обязательных элементов, он не загружает надстройку. Например, если указана версия набора обязательных элементов 1.3, надстройка не отобразится в каком-либо клиенте Outlook, который не поддерживает версии 1.3. и ниже

> [!NOTE]
> Чтобы использовать API в любом из нумерованных наборов обязательных элементов, следует ссылаться на **рабочую** библиотеку в сети CDN (https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> Сведения об использовании API предварительных версий см. в разделе [Использование предварительных версий API](#using-preview-apis) далее в этой статье.

## <a name="using-apis-from-later-requirement-sets"></a>Использование API из наборов обязательных элементов более поздних версий

Установка набора обязательных элементов не ограничивает доступные API, которые может использовать надстройка. Например, если для надстройки указан набор обязательных элементов "Mailbox 1.1", но она выполняется в клиенте Outlook, который поддерживает версию "Mailbox 1.3", надстройка может использовать API из набора обязательных элементов "Mailbox 1.3".

Чтобы использовать более новые API, разработчики могут проверить, поддерживает ли конкретное приложение набор обязательных элементов, выполнив следующее.

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Кроме того, разработчики могут проверить наличие более новых API с помощью стандартных методов JavaScript.

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

Такие проверки не нужно выполнять для API-интерфейсов, присутствующих в версии набора обязательных элементов, указанной в манифесте.

## <a name="choosing-a-minimum-requirement-set"></a>Выбор минимального набора обязательных элементов

Разработчикам следует использовать набор обязательных элементов самой ранней версии, содержащий набор критически важных API для сценария их работы, без которого надстройка не будет работать.

## <a name="requirement-sets-supported-by-exchange-servers-and-outlook-clients"></a>Наборы обязательных элементов, поддерживаемые серверами Exchange и клиентами Outlook

В этом разделе указан диапазон наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook. Сведения о требованиях к серверу и клиенту для запуска надстроек Outlook см. в статье [Требования надстроек Outlook](../../outlook/add-in-requirements.md).

> [!IMPORTANT]
> Если целевой сервер Exchange и клиент Outlook поддерживают разные наборы обязательных элементов, вы ограничены применением более низкой версии набора обязательных элементов. Например, если надстройка работает в Outlook 2016 для Mac (максимальная версия набора обязательных элементов: 1.6) с использованием Exchange 2013 (максимальная версия набора обязательных элементов: 1.1), ваша надстройка ограничивается применением набора обязательных элементов 1.1.

### <a name="exchange-server-support"></a>Поддержка сервера Exchange

Указанные ниже серверы поддерживают надстройки Outlook.

| Продукт | Основная версия Exchange | Поддерживаемые наборы обязательных элементов API |
|---|---|---|
| Exchange Online | Последняя сборка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)\* |
| Локальная среда Exchange | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

> [!NOTE]
> \* Чтобы потребовать набор API удостоверений 1.3 в коде надстройки, проверьте, поддерживается ли он вызовом `isSetSupported('IdentityAPI', '1.3')`. Объявление в манифесте надстройки не поддерживается. Также можно определить, поддерживается ли API, проверив, не `undefined` ли он. Подробнее см. в статье [Использование API из наборов требования более поздних версий](#using-apis-from-later-requirement-sets).

### <a name="outlook-client-support"></a>Поддержка клиентов Outlook

Надстройки поддерживаются в Outlook на следующих платформах.

| Платформа | Основная версия Office или Outlook | Поддерживаемые наборы обязательных элементов API |
|---|---|---|
| Windows | Подписка на Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>, [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup>, [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<sup>1</sup><br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 2019 одноразовая покупка (розничная торговля) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| 2019 одноразовая покупка (корпоративная лицензия) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2016 одноразовая покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
|| 2013 одноразовая покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>3</sup>, [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
| Mac | текущий пользовательский интерфейс<br>(подключено к подписке на Microsoft 365) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| новый пользовательский интерфейс (предварительная версия)<sup>4</sup><br>(подключено к подписке на Microsoft 365) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 2019 одноразовая покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2016 одноразовая покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Подписка на Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Android | Подписка на Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Веб-браузер | современный пользовательский интерфейс Outlook при подключении к<br>Exchange Online: подписка Microsoft 365, Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| классический пользовательский интерфейс Outlook при подключении к<br>Локальная среда Exchange | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup> Поддержка версии **1.8** в Outlook для Windows с подпиской на Microsoft 365 или розничной версии для единовременной покупки доступна начиная с версии 1910 (сборка 12130.20272). Поддержка версии **1.9** в Outlook для Windows с подпиской на Microsoft 365 доступна начиная с версии 2008 (сборка 13127.20296). Поддержка версии **1.10** в Outlook для Windows с подпиской на Microsoft 365 доступна начиная с версии 2104 (сборка 13929.20296). Дополнительные сведения относительно вашей версии см.в журнале обновлений на стр [Office 2019](/officeupdates/update-history-office-2019) или [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) и в статье [Поиск версии клиента Office и канала обновления](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> <sup>2</sup>Чтобы потребовать набор API удостоверений 1.3 в коде надстройки, проверьте, поддерживается ли он вызовом `isSetSupported('IdentityAPI', '1.3')`. Объявление в манифесте надстройки не поддерживается. Также можно определить, поддерживается ли API, проверив, не `undefined` ли он. Подробнее см. в статье [Использование API из наборов требования более поздних версий](#using-apis-from-later-requirement-sets).
>
> <sup>3</sup>Поддержка версии 1.3 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3114349) от 8 декабря 2015 г.](https://support.microsoft.com/kb/3114349) Поддержка версии 1.4 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3118280) от 13 сентября 2016 г.](https://support.microsoft.com/help/3118280) Поддержка версии 1.4 в Outlook 2016 (единовременная покупка) добавлена в рамках [обновления для Office 2016 (KB4022223) от 3 июля 2018 г.](https://support.microsoft.com/help/4022223).
>
> <sup>4</sup> Поддержка нового пользовательского интерфейса Mac (предварительной версии) доступна в Outlook с версии 16.38.506. Дополнительные сведения см. в разделе [Поддержка надстроек в Outlook в новом интерфейсе Mac](../../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).
>
> <sup>5</sup> В настоящее время при проектировании и внедрении надстроек для мобильных клиентов следует учитывать и другие факторы. Например, единственный поддерживаемый режим — это "Сообщение прочитано". Дополнительные сведения см. в статье [Рекомендации по использованию кода при добавлении поддержки для команд надстроек Outlook Mobile](../../outlook/add-mobile-support.md#code-considerations).

> [!TIP]
> Классическую и современную версии Outlook в веб-браузере можно различить по внешнему виду панели инструментов почтового ящика.
>
> **современная версия**
>
> ![Снимок части экрана с изображением панели инструментов современной версии Outlook.](../../images/outlook-on-the-web-new-toolbar.png)
>
> **классическая версия**
>
> ![Снимок части экрана с изображением панели инструментов классической версии Outlook.](../../images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>Использование предварительных версий API

Новые API JavaScript для Outlook сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей. Чтобы отправить отзыв о предварительной версии API, используйте способ обратной связи, представленный в конце веб-страницы с описанием API.

> [!NOTE]
> API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде.

Дополнительные сведения о предварительных версиях интерфейсов API см. в статье [Предварительная версия набора обязательных элементов API для Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
