# <a name="add-in-commands-requirement-sets"></a>Наборы требований для команд надстроек

Наборы требований  — это именованные группы обязательных элементов API. С помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Дополнительные сведения см. в статьях [Команды надстроек для Excel, Word и PowerPoint](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands) и [Команды надстроек Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).

У первого выпуска команд надстроек нет соответствующего набора требований (то есть набора требований AddInCommands 1.0 не существует). В приведенной ниже таблице указаны ведущие приложения Office, которые поддерживают первый выпуск, а также их номера версии или сборки.  

| Выпуск   |  Office 2013 для Windows | 2016 Office для Windows (без подписки) | Office 365 для Windows   |  Office для iPad  |  Office 365 для Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Команды надстроек (первый выпуск, без набора требований) | Н/Д | 16.0.4678.1000 *поддерживается только для Outlook* |Версия 1603 (сборка 6769.0000) или более поздняя | Н/Д | 15.33 или более поздняя версия| Январь 2016 г. | |

В наборе  требований команд надстроек версии 1.1 появилась возможность [автоматического открытия области задач с документами](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

В приведенной ниже таблице указаны наборы требований команд версии 1.1, ведущие приложения Office, которые их поддерживают, и их номера версии или сборки ведущего приложения Office. 

|  Набор требований  |  Office 2013 для Windows | 2016 Office для Windows (без подписки) | Office 365 для Windows   |  Office для iPad  |  Office 365 для Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | Н/Д | 16.0.4678.1000 *поддерживается только для Outlook*  | Версия 1705 (сборка 8121.1000) или более поздняя | Н/Д | 15.34 или более поздняя версия| Май 2017 г. | |

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Какую версию Office Я использую?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы требований API для Office

Сведения об общих наборах требований API см. в статье [Общие наборы требований API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
