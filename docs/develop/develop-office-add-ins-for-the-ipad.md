---
title: Разработка надстроек Office для iPad
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 3ac8f651ccb87b32679a28684f0d08fad53aa773
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952091"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>Разработка надстроек Office для iPad


В приведенной ниже таблице перечислены действия по созданию надстройки Office, которая будет работать в Office для iPad.


|**Задача**|**Описание**|**Ресурсы**|
|:-----|:-----|:-----|
|Обновление надстройки для поддержки Office.js версии 1.1.|Обновите до версии 1.1. файлы JavaScript (Office.js и JS-файлы приложения) и файл проверки манифеста надстройки, которые используете в проекте надстройки Office.|[Изменения API JavaScript для Office](/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office)|
|Следуйте рекомендациям по оформлению пользовательского интерфейса.|Органично интегрируйте в iOS пользовательский интерфейс надстройки.|[Разработка для iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Следуйте рекомендациям по оформлению надстройки.|Убедитесь, что ваша надстройка интересная, полезная и стабильно работает.|[Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md)|
|Оптимизируйте надстройку для сенсорного ввода.|Сделайте так, чтобы пользовательский интерфейс поддерживал не только клавиатуру и мышь, но и сенсорный ввод.|[Принципы разработки пользовательского интерфейса](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Сделайте надстройку бесплатной.|Office на iPad — это канал, через который вы можете привлекать пользователей и рекламировать свои службы. Эти пользователи могут стать вашими клиентами.|[Политика проверки 10.8](/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|Сделайте надстройку некоммерческой.|У надстройки не должно быть пробных версий, она не должна содержать платных возможностей, рекламы платных версий или ссылок на интернет-магазины, в которых пользователи могут приобрести другой контент, приложения или надстройки. На страницах с политикой конфиденциальности и условиями использования также не должно быть рекламы и ссылок на AppSource.|[Политика проверки 3.4](/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|Отправьте свою надстройку в AppSource еще раз.|В службе "Панель мониторинга продаж" установите флажок **Включить эту надстройку в каталог надстроек Office для iPad** и укажите свой идентификатор разработчика Apple в поле "Идентификатор Apple ID". Просмотрите [соглашение с поставщиком приложений](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm).|[Сделайте свои решения доступными в AppSource и Office](/office/dev/store/submit-to-the-office-store)|

Для других платформ надстройку Office можно оставить без изменений. Кроме того, у надстройки может быть различный интерфейс в зависимости от браузера или устройства. Чтобы определить, запущена ли надстройка на iPad, можно использовать следующие API:
- var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)


## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Рекомендации по разработке надстроек Office для iOS и Mac

Следуйте приведенным ниже рекомендациям по разработке надстроек для iOS.


-  **Создайте надстройку с помощью Visual Studio.**

    При разработке надстройки с помощью Visual Studio можно [задать точки останова и выполнить отладку кода](../develop/create-and-debug-office-add-ins-in-visual-studio.md) в ведущем приложении Office на устройстве с Windows, прежде чем загружать неопубликованную надстройку на iPad или Mac. Так как надстройка, работающая в Office для iOS или Office для Mac, поддерживает те же API, что и надстройка в Office в Windows, код надстройки должен выполняться одинаково на обеих платформах.

-  **Укажите требования касательно API в манифесте надстройки или с помощью проверок в среде выполнения.**

    Когда вы укажете требования к API в манифесте надстройки, Office определит, поддерживает ли ведущее приложение эти элементы API. Если нужные элементы API доступны в ведущем приложении, то надстройка будет доступна в нем. Кроме того, вы можете выполнить проверку в среде выполнения, чтобы определить, доступен ли метод в ведущем приложении, прежде чем использовать его в надстройке. Проверки в среде выполнения гарантируют постоянную доступность самой надстройки в ведущем приложении, а также при наличии соответствующих методов — ее дополнительных функций. Дополнительные сведения см. в статье [Указание ведущих приложений Office и требований API](specify-office-hosts-and-api-requirements.md).

Общие рекомендации по разработке надстроек см. в статье [Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md).


## <a name="see-also"></a>См. также

- [Загрузка неопубликованной надстройки Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Отладка надстроек Office на iPad и Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
