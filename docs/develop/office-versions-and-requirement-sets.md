---
title: Версии Office и наборы требований
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: ac3ae4fa3eeca9cfbd56b15168fc39d67139680d
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505995"
---
# <a name="office-versions-and-requirement-sets"></a>Версии Office и наборы требований

Имеется множество версий Office на нескольких платформах, и они не поддерживают каждый API в JavaScript API для Office  (Office.js). Вы не всегда можете контролировать версию Office, которую установили ваши пользователи. Чтобы справиться с этой ситуацией, мы предоставляем систему, называемую наборами требований, которые помогут вам определить, поддерживает ли ведущее приложение Office возможности, необходимые в надстройке для вашего Office. 

> [!NOTE]
> - Office работает на разных платформах, в том числе Office для Windows, Office Online, Office для Mac и Office для iPad.  
> - Примеры ведущих приложений Office — Excel, Word, PowerPoint, Outlook, OneNote и другие продукты.  
> - Набор требований — это именованная группа элементов API, например, `ExcelApi 1.5`, `WordApi 1.3` и т. д.  


## <a name="how-to-check-your-office-version"></a>Как узнать, какая версия Office используется

Чтобы определить версию Office, которую вы используете, из приложения Office выберите меню **Файл**, а затем выберите **Аккаунт**. Версия Office появится в разделе ** Информация о продукте**. Например, следующий снимок экрана указывает на версию Office 1802 (сборка 9026.1000):

![Проверка версии Office](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Доступность наборов требований для Office

Надстройки Office могут использовать наборы требований API, чтобы определить, поддерживает ли ведущее приложение Office членов API, которые оно должен использовать. Поддержка набора требований зависит от ведущего приложения Office и версии ведущего приложения Office (см. предыдущий раздел).

Некоторые ведущие приложения Office имеют свои собственные наборы требований API. Например, первым требованием, установленным для API Excel, было `ExcelApi 1.1`, а первым требованием, установленным для API Word, было `WordApi 1.1`. С тех пор было добавлено несколько новых наборов требований ExcelApi и наборов требований WordApi для обеспечения дополнительных функциональных возможностей API.

Кроме того, к общему API добавлены другие функции, такие как команды надстройки (расширение ленты) и возможность запуска диалоговых окон (Dialog API). Команды надстройки и наборы требований API Dialog являются примерами наборов API, которые совместно используют различные ведущие приложения Office.

Надстройка может использовать API только в наборах требований, которые поддерживаются версией ведущего приложения Office, где работает надстройка. Чтобы точно знать, какие из наборов требований доступны для конкретной версии ведущего приложения Office, обратитесь к следующим наборам требований к конкретному ведущему приложению:

- [Наборы требований API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)
- [Наборы требований API JavaScript для Word](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)
- [Наборы требований API JavaScript для OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)
- [Общие сведения о наборах требований API Outlook](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)

Некоторые наборы требований содержат API, которые могуть использоваться любым ведущим приложением Office. Для информации об этих наборах требований см. следующие статьи:

- [Общие наборы требований для Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [Наборы требований для команд надстроек](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [Наборы требований API общих диалогов](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Наборы требований API удостоверений](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

Номер версии набора требований, например "1.1" в `ExcelApi 1.1`, относится к ведущему приложению Office. Номер версии заданного набора требований (например, `ExcelApi 1.1`) не соответствует номеру версии Office.js или наборам требований для других ведущих приложений Office (например, Word, Outlook и т. д.). Наборы требований для разных ведущих приложений Office выпускаются с разной скоростью и временем. Например, `ExcelApi 1.5` был выпущен до установки требования `WordApi 1.3`.

Библиотека API JavaScript для Office (Office.js) включает в себя все наборы требований, которые в настоящее время доступны. В то время как имеются наборы требований `ExcelApi 1.3` и `WordApi 1.3`, отсутствует набор требований `Office.js 1.3`. Последняя версия Office.js поддерживается как одна конечная точка Office, поставляемая через сеть доставки контента (CDN). Для получения дополнительной информации о CDN Office.js, в том числе о том, как управлять версиями и обратной совместимостью, см. [Понимание JavaScript API для Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

## <a name="specify-office-hosts-and-requirement-sets"></a>Указание ведущих приложений Office и наборов требований

Существуют различные способы указать, какие ведущие приложения Office и наборы требований требуются для надстройки. Для получения подробной информации см. [Задание ведущих приложений и требования к API. ](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)


## <a name="see-also"></a>См. также

- [Указание ведущих приложений Office и требований API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Установка последней версии Office](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [Обзор каналов обновления Office 365 профессиональный плюс](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Получите максимум от наших продуктов благодаря Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
