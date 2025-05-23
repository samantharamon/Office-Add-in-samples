---
page_type: sample
products:
  - m365
  - office
  - office-outlook
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
description: "Сведения о создании надстройки Microsoft Outlook, подключающейся к Microsoft Graph"
---

# Получение книг Excel с помощью Microsoft Graph и MSAL в надстройке Outlook 

Узнайте, как создать надстройку Microsoft Outlook, которая подключается к Microsoft Graph, находит первые три книги, сохраненные в OneDrive для бизнеса, извлекает их имена и вставляет имена в новую форму создания сообщения в Outlook.

## Функции

Интегрируя данные поставщиков интернет-служб, вы повышаете ценность и популярность своих надстроек. В этом примере кода показано, как подключить надстройку Outlook к Microsoft Graph. С его помощью можно:

* подключиться к Microsoft Graph из надстройки Office;
* использовать библиотеку MSAL .NET для внедрения инфраструктуры авторизации OAuth 2.0 в надстройке;
* использовать REST API для OneDrive из Microsoft Graph;
* отображать диалоговое окно с использованием пространства имен пользовательского интерфейса Office;
* создать надстройку с помощью ASP.NET MVC, MSAL 3.x.x для .NET и Office.js. 

## Область применения

-  Outlook на всех платформах

## Предварительные требования

Чтобы запустить этот пример кода, необходимо следующее:

* Visual Studio 2019 или более поздней версии.

* SQL Server Express (если автоматически не установлен с последними версиями Visual Studio).

* Учетная запись Office 365, которую получают участники [программы для разработчиков Office 365](https://aka.ms/devprogramsignup), предоставляющаяся вместе с бесплатной годичной подпиской на Office 365.

* Минимум три книги Excel, сохраненные в OneDrive для бизнеса в составе подписки на Office 365.

* (Необязательно) Если вы хотите выполнить отладку в классической версии, а не в Outlook Online: Outlook для Windows версии 1809 или более поздней.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Клиент Microsoft Azure. Эта надстройка требует наличия Azure Active Directiory (AD). В Azure AD доступны службы идентификации, которые приложения используют для проверки подлинности и авторизации. Здесь можно получить пробную подписку: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Решение

Решение | Авторы
---------|----------
Надстройка Outlook в Microsoft Graph ASP.NET | Майкрософт

## Журнал версий

Версия | Дата | Примечания
---------| -----| --------
1.0 | 8 июля 2019 г. | Первый выпуск

## Заявление об отказе

**ЭТОТ КОД ПРЕДОСТАВЛЯЕТСЯ *КАК ЕСТЬ* БЕЗ КАКОЙ-ЛИБО ЯВНОЙ ИЛИ ПОДРАЗУМЕВАЕМОЙ ГАРАНТИИ, ВКЛЮЧАЯ ПОДРАЗУМЕВАЕМЫЕ ГАРАНТИИ ПРИГОДНОСТИ ДЛЯ КАКОЙ-ЛИБО ЦЕЛИ, ДЛЯ ПРОДАЖИ ИЛИ ГАРАНТИИ ОТСУТСТВИЯ НАРУШЕНИЯ ПРАВ ИНЫХ ПРАВООБЛАДАТЕЛЕЙ.**

----------

## Построение и запуск решения

## Настройка решения

1. В **Visual Studio** выберите проект **Outlook-Add-in-Microsoft-Graph-ASPNETWeb**. Убедитесь, что в окне **Свойства** для параметра **SSL включен** задано значение **True**. Убедитесь, что в поле **URL-адрес SSL** используются доменное имя и номер порта, указанные на следующем этапе.
 
2. Зарегистрируйте свое приложение на [портале управления Azure](https://manage.windowsazure.com). **Войдите в систему, используя учетные данные администратора Office 365, чтобы убедиться, что вы работаете в службе Azure Active Directory, связанной с ними.** Сведения о регистрации приложений см. в статье [Регистрация приложения с помощью платформы удостоверений Майкрософт](https://learn.microsoft.com/graph/auth-register-app-v2). Используйте указанные ниже параметры:

 - URI ПЕРЕНАПРАВЛЕНИЯ: https://localhost:44301/AzureADAuth/Authorize
 - ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ "Учетные записи только в этом каталоге организации"
 - НЕЯВНОЕ ПРЕДОСТАВЛЕНИЕ РАЗРЕШЕНИЯ: Не включайте никакие параметры неявного предоставления разрешений
 - РАЗРЕШЕНИЯ API (делегированные, не разрешения приложений): **Files.Read.All** и **User.Read**

	> Примечание. После регистрации приложения скопируйте **идентификатор приложения (клиента)** и **идентификатор директории (клиента)** в колонке **Обзор** регистрации приложения на портале управления Azure. Также скопируйте секретный код клиента, созданный в колонке **Сертификаты и секреты**. 
	 
3.  В узле web.config используйте значения, скопированные на предыдущем этапе. Для параметра **AAD:ClientID** задайте значение идентификатора клиента, а для параметра **AAD:ClientSecret** — значение секретного кода клиента. Задайте ваш идентификатор клиента Office 365 в **"AAD:O365TenantID"**. 

## Запуск решения

1. Откройте файл решения в Visual Studio. 
2. Щелкните правой кнопкой мыши решение **Outlook-Add-in-Microsoft-Graph-ASPNET** в **Обозревателе решений** (не узлы проекта) и выберите **Назначить запускаемые проекты**. Установите переключатель **Несколько запускаемых проектов**. Убедитесь, что проект, имя которого заканчивается на "Web", указан первым.
3. В меню **Сборка** выберите команду **Очистить решение**. После выполнения команды снова откройте меню **Сборка** и выберите **Собрать решение**.
4. В **Обозревателе решений** выберите узел проекта **Outlook-Add-in-Microsoft-Graph-ASPNET** (не верхний узел решения и не узел проекта, имя которого заканчивается на "Web").
5. В области **Свойства** откройте раскрывающийся список **Действие при запуске** и выберите запуск надстройки в классической версии Outlook или Outlook в Интернете в одном из перечисленных браузеров. (*Не выбирайте Internet Explorer. Причины см. в разделе **Известные проблемы** ниже.*) 

    ![Выберите нужное ведущее приложение Outlook: классическое или в одном из браузеров](images/StartAction.JPG)

6. Нажмите клавишу F5. При первом запуске вам будет предложено указать адрес электронной почты и пароль пользователя, которые будут использоваться для отладки надстройки. Используйте учетные данные администратора своего клиента Office 365. 

    ![Форма с текстовыми полями для электронной почты и пароля пользователя](images/CredentialsPrompt.JPG)

    >ПРИМЕЧАНИЕ. Браузер откроется на странице входа в Office в Интернете. (Если это первый запуск надстройки, вы введете имя пользователя и пароль дважды.) 

Остальные действия зависят от среды работы надстройки: классическая версия Outlook или Outlook в Интернете.

### Запуск решения в Outlook в Интернете

1. Outlook в Интернете откроется в окне браузера. В Outlook щелкните **Создать**, чтобы создать сообщение электронной почты. 
2. Под областью создания сообщения расположена панель инструментов с кнопками **Отправить**, **Отменить** и другими функциями. В зависимости от используемого интерфейса **Outlook в Интернете** значок надстройки расположен с правого края этой панели инструментов или в раскрывающемся меню, появляющемся при нажатии кнопки **...** на этой панели инструментов.

   ![Значок для надстройки "Вставка файлов"](images/Onedrive_Charts_icon_16x16px.png)

3. Щелкните значок, чтобы открыть надстройку области задач.
4. С помощью надстройки добавьте имена первых трех книг из учетной записи OneDrive пользователя в сообщение. Страницы и кнопки надстройки не требуют объяснений.

## Запуск проекта в классической версии Outlook

1. Откроется классическое приложение Outlook. В Outlook щелкните **Создать сообщение**, чтобы создать сообщение электронной почты. 
2. На ленте **Сообщение** формы **Сообщение** есть кнопка **Открыть надстройку** в группе **Файлы OneDrive**. Нажмите кнопку, чтобы открыть надстройку.
3. С помощью надстройки добавьте имена первых трех книг из учетной записи OneDrive пользователя в сообщение. Страницы и кнопки надстройки не требуют объяснений.

## Известные проблемы

* Индикатор работы Fabric отображается кратковременно или совсем не отображается. 
* Если вы работаете в Internet Explorer, при попытке входа возникает ошибка с сообщением о необходимости разместить `https://localhost:44301` и `https://outlook.office.com` (или `https://outlook.office365.com`) в одной зоне безопасности. Но эта ошибка появляется даже после выполнения этого действия. 

## Вопросы и комментарии

Мы будем рады получить ваши отзывы о примере *получения книг Excel с помощью Microsoft Graph и MSAL в надстройке Office*.
Своими мыслями можете поделиться на вкладке *Проблемы* этого репозитория. Общие вопросы о разработке решений для Office 365 следует публиковать на сайте [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Помечайте свои вопросы тегами [office-js], [MicrosoftGraph] и [API].

## Дополнительные ресурсы

* [Документация по Microsoft Graph](https://learn.microsoft.com/graph/)
* [Документация по надстройкам Office](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Авторские права
© Корпорация Майкрософт (Microsoft Corporation), 2019. Все права защищены.

Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Outlook-Add-in-Microsoft-Graph-ASPNET" />
