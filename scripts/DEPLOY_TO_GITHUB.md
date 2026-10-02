# Инструкция по деплою и автономному запуску на GitHub Pages

Данная инструкция описывает процесс публикации генератора 8 XML-фидов Авито компании **«Дарион Свет»** в облако GitHub. 

Генерация и хостинг фидов будут выполняться **бесплатно, автономно и круглосуточно** без необходимости держать включенным рабочий компьютер.

---

## 1. Команды для публикации проекта в GitHub (Git Setup)

Выполните следующие команды в корневой папке проекта (`D:\Automatizations\Fl.ru\Darion свет\`):

```bash
# 1. Инициализация локального репозитория (если еще не инициализирован)
git init

# 2. Добавление всех файлов проекта в индекс и создание первого коммита
git add .
git commit -m "feat: initial commit with Avito feeds generator & GitHub Actions workflow"

# 3. Привязка к вашему удаленному GitHub-репозиторию и отправка ветки main
git branch -M main
git remote add origin https://github.com/<username>/<repo-name>.git
git push -u origin main
```
*(Замените `<username>` на ваш логин GitHub, а `<repo-name>` на имя созданного репозитория).*

---

## 2. Настройка прав GitHub Actions и включение GitHub Pages

### Шаг 2.1. Предоставление прав на запись токену GITHUB_TOKEN:
1. Перейдите в ваш репозиторий на GitHub: **Settings** -> **Actions** -> **General**.
2. Прокрутите страницу вниз до блока **Workflow permissions**.
3. Выберите опцию: **Read and write permissions**.
4. Нажмите **Save**.

### Шаг 2.2. Включение GitHub Pages:
1. Запустите workflow первый раз вручную:
   - Вкладка **Actions** -> выберите workflow **«Update and Deploy Avito XML Feeds»** -> нажмите **Run workflow**.
   - Дождитесь завершения с зеленой галочкой (будет автоматически создана ветка `gh-pages` с содержимым папки `output/`).
2. В репозитории откройте **Settings** -> **Pages**.
3. В разделе **Build and deployment**:
   - **Source**: выберите `Deploy from a branch`.
   - **Branch**: выберите `gh-pages`, папка `/ (root)`.
   - Нажмите **Save**.

---

## 3. Регламент работы автоматизации

- **Расписание (Cron)**: Каждый день в **03:00 UTC** (**06:00 по МСК**) GitHub Actions автоматически:
  1. Разворачивает чистое Linux-окружение (Python 3.11).
  2. Запрашивает свежие остатки и цены из Google Таблицы.
  3. Генерирует все 8 валидных XML-фидов.
  4. Обновляет ветку `gh-pages` без остановки раздачи файлов.
- **Ручной запуск**: В любой момент через вкладку **Actions** -> **Run workflow**.

---

## 4. Итоговые публичные ссылки для кабинета «Автозагрузка Авито»

Укажите эти 8 постоянных ссылок в личном кабинете Авито (**Для бизнеса -> Автозагрузка -> Настройки -> Ссылка на файл**):

1. **LED Светильники (промышленные, уличные, универсальные):**
   `https://<username>.github.io/<repo-name>/led_luminaires_feed.xml`

2. **Трековые системы освещения:**
   `https://<username>.github.io/<repo-name>/track_systems_feed.xml`

3. **Люстры и потолочные светильники:**
   `https://<username>.github.io/<repo-name>/chandeliers_feed.xml`

4. **Лампы и лампочки:**
   `https://<username>.github.io/<repo-name>/lamps_feed.xml`

5. **Светильники общего назначения (Бра, Споты):**
   `https://<username>.github.io/<repo-name>/lighting_fixtures_feed.xml`

6. **Ламповые светильники:**
   `https://<username>.github.io/<repo-name>/lamp_luminaires_feed.xml`

7. **Электрика и комплектующие:**
   `https://<username>.github.io/<repo-name>/electrics_feed.xml`

8. **Прочие товары / Торговое оборудование:**
   `https://<username>.github.io/<repo-name>/other_goods_feed.xml`

---
*Все ссылки работают по безопасному протоколу HTTPS, имеют встроенный CDN от GitHub и доступны для робота-парсера Авито 24/7.*
