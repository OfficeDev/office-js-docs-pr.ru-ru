---
title: Тестирование и отладка надстроек Office
description: Узнайте, как тестировать и отлаживать свою надстройку Office
ms.date: 09/24/2021
ms.localizationpriority: high
ms.openlocfilehash: db0edec5c7b7c741425a9d27d7580a52d2839546
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081416"
---
# <a name="test-and-debug-office-add-ins"></a>Тестирование и отладка надстроек Office

Эта статья содержит рекомендации по тестированию, отладке и диагностике надстроек Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Тестирование кроссплатформенных выпусков и выпусков для нескольких версий Office

Надстройки Office запускаются на основных платформах, поэтому требуется протестировать надстройку на всех платформах, где ваши пользователи могут запускать Office. К ним обычно относятся Office в Интернете, Office для Windows (как подписка, так и единовременная покупка), Office для Mac, Office для iOS и (для надстроек Outlook) Office для Android. Однако могут возникать ситуации, когда вы точно знаете, что никто из ваших пользователей не будет работать на некоторых платформах. Например, если вы создаете надстройку для компании, которая требует, чтобы пользователи работали на компьютерах с Windows и подпиской на Office, вам не нужно выполнять тестирование в Office для Mac или единовременно приобретенных экземплярах для Windows.

> [!NOTE]
> На компьютерах с Windows браузер, используемый надстройкой, определяется версией Windows и Office. Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Надстройки, предлагаемые через AppSource, проходят проверку, включающую тестирование на всех платформах. Кроме того, надстройки тестируются в Office для Интернета со всеми основными современными браузерами, включая Microsoft Edge (WebView2 на основе Chromium), Chrome и Safari. Соответственно, перед отправкой в AppSource необходимо протестировать эти платформы и браузеры. Дополнительные сведения о проверке см. в статье [Политики сертификации коммерческой платформы Marketplace](/legal/marketplace/certification-policies), особенно в [разделе 1120.3](/legal/marketplace/certification-policies#11203-functionality), а также на странице [Доступность и применение надстроек Office](../overview/office-add-in-availability.md).
>
> AppSource не использует Internet Explorer или устаревшую версию Microsoft Edge (WebView1) для тестирования надстроек в Office для Интернета. Но если значительное число ваших пользователей будет использовать браузер Edge прежних версий для открытия Office в Интернете, вам следует протестировать надстройку с ним. (Office в Интернете не будет открываться в Internet Explorer, поэтому тестировать надстройку с этим браузером не нужно.) Дополнительные сведения см. в статьях "[Поддержка Internet Explorer 11](../develop/support-ie-11.md)" и "[Устранение неполадок Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)". Office по-прежнему поддерживает эти браузеры как поставщики сред выполнения надстроек, поэтому если вы считаете, что столкнулись с ошибкой в работе надстроек в них, создайте запись о проблеме для репозитория [office-js.](https://github.com/OfficeDev/office-js/issues/new/choose)

## <a name="sideload-an-office-add-in-for-testing"></a>Загрузка неопубликованной надстройки Office для тестирования

Вы можете установить надстройку Office для тестирования, не размещая ее в каталоге надстроек. Процедура отличается для разных платформ, а в некоторых случаях и для разных продуктов. Следующие статьи посвящены загрузке неопубликованных надстроек Office на определенной платформе или в определенном продукте.

- [Загрузка неопубликованных надстроек Office в Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Загрузка неопубликованных надстроек Office в Office в Интернете](sideload-office-add-ins-for-testing.md)

- [Загрузка неопубликованных надстроек Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

- [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>Отладка надстройки Office

Процедура отладки также отличается для разных платформ. Следующие статьи посвящены отладке надстроек Office на определенной платформе.

- [Подключение отладчика из области задач (в Windows)](attach-debugger-from-task-pane.md)
- [Отладка надстроек с помощью средств разработчика для Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Отладка надстроек с помощью средств разработчика для устаревшей версии Microsoft Edge](debug-add-ins-using-devtools-edge-legacy.md)
- [Отладка надстроек с помощью средств разработчика в Microsoft Edge (на основе Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md)
- [Отладка надстроек Office на Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>Проверка манифеста надстройки Office

Информацию о проверке манифеста надстройки Office и устранении связанных с ним неполадок см. в [этой статье](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Устранение ошибок, с которыми сталкиваются пользователи

Информацию об устранении основных ошибок, с которыми сталкиваются пользователи при работе с надстройками Office, см. в [этой статье](testing-and-troubleshooting.md).
