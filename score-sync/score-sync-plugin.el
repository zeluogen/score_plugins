;; score-sync-plugin.el - 学生成绩批量输入插件
;; 编辑器版本：Emacs28.2
;; 插件目录：~/.emacs.d/score-sync/
;; 分组文件目录：~/.emacs.d/score-sync/groups/
;; 快捷键：C-c s s - 执行成绩同步
;;         C-c s g - 打开分组文件夹

;; 关闭字节编译警告
(setq byte-compile-warnings '(not free-vars unresolved noruntime lexical make-local))

;; -------------------------- 全局变量定义 --------------------------
;; 插件根目录
(defvar score-sync-root-dir (expand-file-name "~/.emacs.d/score-sync/")
  "插件根目录")

;; 分组文件存放目录
(defvar score-sync-groups-dir (expand-file-name "groups" score-sync-root-dir)
  "分组文件存放目录")

;; 评分标准文件存放目录
(defvar score-sync-rubrics-dir (expand-file-name "rubrics" score-sync-root-dir)
  "评分标准文件存放目录")

;; 小组映射哈希表：键=学号(字符串)，值=同组所有学号列表(字符串列表)
(defvar score-sync-group-map (make-hash-table :test 'equal)
  "学号到同组学号的映射表")

;; 成绩列名标识
(defvar score-sync-score-col-name "score"
  "成绩列名标识，用于识别成绩列位置")

;; 学号正则表达式（匹配12位数字学号）
(defvar score-sync-student-no-regex "[0-9]\\{12\\}"
  "匹配学生学号的正则表达式")

;; -------------------------- 基础工具函数 --------------------------
(defun score-sync-make-dirs ()
  "创建插件所需目录（根目录+分组目录+评分标准目录），若已存在则不操作"
  (interactive)
  (dolist (dir (list score-sync-root-dir score-sync-groups-dir score-sync-rubrics-dir))
    (unless (file-directory-p dir)
      (message "[score-sync] 创建目录：%s" dir)
      (make-directory dir t)
      (message "[score-sync] 成功创建目录：%s" dir))))

(defun score-sync-split-line-by-space (line)
  "按任意空白符分割行内容，忽略分隔符数量，返回列内容列表
  ARG line: 待分割的行字符串"
  (split-string line "\\s-+" t))  ; \\s-+匹配任意空白符，t忽略空元素

(defun score-sync-find-score-col-index ()
  "从成绩文件中找到SCORE列的索引（从0开始）：
  1. 找到以#开头的列名行
  2. 按空白符分割列名，匹配score列的位置
  报错：未找到列名行、列名行无score列"
  (save-excursion
    (goto-char (point-min))  ; 跳转到文件开头
    (let ((col-index nil))
      ;; 遍历行找列名行：以#开头，包含student.no/student.name/score核心列
      (while (and (not col-index) (not (eobp)))
        (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                            (line-end-position))))
          (when (and (string-prefix-p "#" current-line)
                     (string-match-p "student\\.no" current-line)
                     (string-match-p "student\\.name" current-line))
            ;; 分割列名并查找score列索引
            (let ((col-names (score-sync-split-line-by-space (substring current-line 1))))
              (setq col-index (cl-position score-sync-score-col-name col-names :test 'equal))))
          (forward-line 1)))
      ;; 报错检查
      (unless col-index
        (error "[score-sync] 错误：未找到成绩列（%s），请检查成绩文件列名行格式" score-sync-score-col-name))
      col-index)))

(defun score-sync-get-current-student-no ()
  "提取光标所在行的学生学号，返回学号字符串
  报错：当前行无内容、第一列非有效学号"
  (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                      (line-end-position))))
    (if (string-blank-p current-line)
        (error "[score-sync] 错误：光标所在行为空，无学生信息")
      (let ((col-list (score-sync-split-line-by-space current-line)))
        (if (null col-list)
            (error "[score-sync] 错误：光标所在行无有效列内容")
          (let ((first-col (car col-list)))
            (if (string-match score-sync-student-no-regex first-col) 
                (match-string 0 first-col) 
              (error "[score-sync] 错误：第一列<%s>非有效学号" first-col))))))))

(defun score-sync-get-current-score (score-col-index)
  "提取光标所在行的成绩内容，返回成绩字符串（空则返回空字符串）
  若列数不足成绩列索引，提醒用户添加分隔符后再输入成绩
  ARG score-col-index: 成绩列索引（从0开始）"
  (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                      (line-end-position)))
        (score ""))
    (let ((col-list (score-sync-split-line-by-space current-line)))
      (if (< (length col-list) (1+ score-col-index))
          ;; 列数不足，提醒添加分隔符
          (warn "[score-sync] 提醒：当前行列数不足，需先添加分隔符再输入成绩！")
        ;; 提取成绩列内容
        (setq score (nth score-col-index col-list))))
    score))

(defun score-sync-find-student-line (student-no)
  "在当前缓冲区中查找指定学号的行，返回行起始位置
  报错：未找到对应学号的行
  ARG student-no: 待查找的学号字符串"
  (save-excursion
    (goto-char (point-min))
    (let ((target-pos nil))
      (while (and (not target-pos) (not (eobp)))
        (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                            (line-end-position))))
          (when (string-prefix-p student-no current-line)
            (setq target-pos (line-beginning-position))))
        (forward-line 1)))
    (unless target-pos
      (error "[score-sync] 错误：未找到学号<%s>对应的行" student-no))
    target-pos))

;; -------------------------- 成绩预输入相关函数 --------------------------
(defun score-sync-get-score-file-title ()
  "获取当前成绩文件的标题，返回标题字符串
  标题为文件中第一个以#开头的行，去除#和空白符"
  (save-excursion
    (goto-char (point-min))
    (let ((title nil))
      (while (and (not title) (not (eobp)))
        (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                            (line-end-position))))
          (when (string-prefix-p "#" current-line)
            (setq title (string-trim (substring current-line 1))))
          (forward-line 1)))
      title)))

(defun score-sync-parse-rubric-file (file-path)
  "解析单个评分标准文件，返回包含标题、学号列表和总分数的列表
  返回格式：(title student-no-list total-score)
  ARG file-path: 评分标准文件的绝对路径"
  (when (file-readable-p file-path)
    (with-temp-buffer
      (insert-file-contents file-path)
      (let ((title nil)
            (student-no-list nil)
            (total-score nil))
        ;; 查找标题行（#开头）
        (goto-char (point-min))
        (while (and (not title) (not (eobp)))
          (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                              (line-end-position))))
            (when (string-prefix-p "#" current-line)
              (setq title (string-trim (substring current-line 1))))
            (forward-line 1)))
        ;; 查找Total行和提取学号
        (goto-char (point-min))
        (while (not (eobp))
          (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                              (line-end-position))))
            ;; 提取学号
            (when (string-match score-sync-student-no-regex current-line)
              (let ((no (match-string 0 current-line)))
                (add-to-list 'student-no-list no :test 'equal)))
            ;; 提取Total分数
            (when (string-match "\\*Total\\*.*--\\s-*\\([0-9]+\\)/[0-9]+" current-line)
              (setq total-score (match-string 1 current-line)))
            (forward-line 1)))
        (list title student-no-list total-score)))))

(defun score-sync-auto-fill-scores ()
  "自动填充成绩函数
  1. 获取当前成绩文件标题
  2. 遍历rubrics目录下的文件，找到标题匹配的文件
  3. 解析匹配文件，提取学号列表和总分数
  4. 检查学号是否在分组映射表中存在
  5. 填充成绩到对应学生行"
  (interactive)
  ;; 创建插件目录
  (score-sync-make-dirs)
  ;; 加载分组映射（无论哈希表是否为空）
  (score-sync-load-all-groups)
  ;; 获取成绩文件标题
  (let ((score-file-title (score-sync-get-score-file-title)))
    (unless score-file-title
      (error "[score-sync] 错误：未找到成绩文件标题（#标记）"))
    (message "[score-sync] 正在预输入成绩，文件标题：%s" score-file-title)
    ;; 确保rubrics目录存在
    (let ((rubrics-dir-exists (file-directory-p score-sync-rubrics-dir)))
      (message "[score-sync] 检查rubrics目录：%s，存在：%s" score-sync-rubrics-dir rubrics-dir-exists)
      (unless rubrics-dir-exists
        (message "[score-sync] 创建rubrics目录：%s" score-sync-rubrics-dir)
        (make-directory score-sync-rubrics-dir t)
        (setq rubrics-dir-exists t)
        (message "[score-sync] 已创建评分标准目录：%s" score-sync-rubrics-dir))
      
      ;; 遍历rubrics目录下的文件
      (if rubrics-dir-exists 
          (let ((rubric-files (directory-files score-sync-rubrics-dir t nil nil))
                (fill-count 0))
            (message "[score-sync] 评分标准目录：%s，文件数：%d" score-sync-rubrics-dir (- (length rubric-files) 2)) ;; 减去.和..目录
            (dolist (file rubric-files)
              (when (and (file-regular-p file) (not (string-prefix-p "." (file-name-nondirectory file))))
                (message "[score-sync] 检查评分标准文件：%s" (file-name-nondirectory file))
                (let ((rubric-data (score-sync-parse-rubric-file file)))
                  (message "[score-sync] 解析结果：%s" rubric-data)
                  (when rubric-data
                    (let ((rubric-title (car rubric-data))
                          (student-no-list (cadr rubric-data))
                          (total-score (caddr rubric-data)))
                      ;; 检查标题是否匹配（忽略大小写）
                      (when (and rubric-title (string-equal (downcase rubric-title) (downcase score-file-title)) total-score)
                        (message "[score-sync] 找到匹配的评分标准文件：%s，总分数：%s" 
                                 (file-name-nondirectory file) total-score)
                        ;; 检查学号是否在分组映射表中存在
                        (let ((valid-students t)
                              (missing-students nil))
                          (dolist (student-no student-no-list)
                            (unless (gethash student-no score-sync-group-map)
                              (setq valid-students nil)
                              (add-to-list 'missing-students student-no)))
                          (if valid-students
                              (progn
                                ;; 获取成绩列索引
                                (let ((score-col-index (score-sync-find-score-col-index)))
                                  ;; 填充成绩
                                  (save-excursion
                                    (goto-char (point-min))
                                    (while (not (eobp))
                                      (let ((current-pos (point)))
                                        (let* ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                                                            (line-end-position)))
                                               (col-list (score-sync-split-line-by-space current-line)))
                                          ;; 处理内容行，跳过标题行（#开头）
                                          (when (and col-list (not (string-blank-p current-line)) (not (string-prefix-p "#" current-line)))
                                            (let ((line-no (car col-list)))
                                              ;; 判断当前行学号是否在rubric文件的学号列表中
                                              (when (member line-no student-no-list)
                                                ;; 列数不足时补充分隔符到成绩列索引
                                                (let ((new-col-list (if (< (length col-list) (1+ score-col-index))
                                                                        (append col-list
                                                                                (make-list (- (1+ score-col-index) (length col-list)) ""))
                                                                      col-list)))
                                                  ;; 替换成绩列内容
                                                  (setf (nth score-col-index new-col-list) total-score)
                                                  ;; 用制表符分隔并重新拼接行
                                                  (let ((new-line (string-join new-col-list "\t")))
                                                    (message "[score-sync] 填充成绩：%s -> %s" line-no total-score)
                                                    ;; 删除原行，插入新行
                                                    (kill-whole-line)
                                                    (insert new-line "\n")
                                                    (setq fill-count (1+ fill-count)))))))
                                      ;; 移动到下一行
                                      (goto-char current-pos)
                                      (forward-line 1)))))
                            (when missing-students
                              (warn "[score-sync] 警告：部分学号在分组映射表中不存在：%s，跳过填充" missing-students))))))))
            (if (> fill-count 0)
                (message "[score-sync] 成绩预输入共填充 %d 名学生" fill-count)
              (message "[score-sync] 未找到匹配的评分标准文件或无有效学号")))))))
        (warn "[score-sync] 警告：评分标准目录<%s>不存在或不可访问" score-sync-rubrics-dir)))))

;; -------------------------- 分组文件解析函数 --------------------------
(defun score-sync-parse-group-file (file-path)
  "解析单个分组文件，提取所有有效学号，返回学号（字符串）列表
  按正则匹配文件中所有12位数字学号，自动去重
  报错：文件无有效学号
  ARG file-path: 分组文件的绝对路径"
  (when (file-readable-p file-path)
    (with-temp-buffer
      (insert-file-contents file-path)  ; 读取文件内容到临时缓冲区
      (let ((content (buffer-string))
            (student-no-list nil))
        ;; 正则匹配有效学号，存入列表
        (while (string-match score-sync-student-no-regex content)
          (let ((no (match-string 0 content)))
            (add-to-list 'student-no-list no :test 'equal))  ; 去重
          (setq content (substring content (match-end 0))))
        ;; 报错：文件无有效学号
        (if (null student-no-list)
            (warn "[score-sync] 警告：分组文件<%s>无有效学号，已跳过" (file-name-nondirectory file-path))
          student-no-list)))))

(defun score-sync-load-all-groups ()
  "加载分组目录下所有文件的小组信息，构建学号到同组学号的映射表
  遍历分组目录，解析每个文件，为每个学号绑定同组所有学号
  自动刷新哈希表，避免重复加载"
  (interactive)
  ;; 清空原有映射，避免重复
  (clrhash score-sync-group-map)
  (if (file-directory-p score-sync-groups-dir)
      (let ((group-files (directory-files score-sync-groups-dir t nil nil)))
        (message "[score-sync] 分组文件数量：%d" (- (length group-files) 2)) ;; 减去.和..目录
        (dolist (file group-files)
          ;; 仅处理普通文件，跳过目录或隐藏文件
          (when (and (file-regular-p file) (not (string-prefix-p "." (file-name-nondirectory file))))
            (let ((student-no-list (score-sync-parse-group-file file)))
              (when student-no-list
                ;; 为学号添加同组映射，同步时跳过自身
                (dolist (no student-no-list)
                  (puthash no student-no-list score-sync-group-map))))))
        (message "[score-sync] 成功加载分组文件，共解析%d个学号映射" (hash-table-count score-sync-group-map))
        (maphash (lambda (key value) (message "学号: %s, 同组学号: %s" key value)) score-sync-group-map))
    (error "[score-sync] 错误：分组目录<%s>不存在或不可访问" score-sync-groups-dir)))

;; -------------------------- 核心同步函数 --------------------------
(defun score-sync-sync-score ()
  "成绩批量输入核心函数，绑定快捷键
  1. 检查并创建插件目录
  2. 加载分组映射表
  3. 识别成绩列索引
  4. 提取当前行学号和成绩
  5. 匹配同组学号，遍历同步成绩
  6. 覆盖同组学生成绩列内容"
  (interactive)
  ;; 创建插件目录
  (score-sync-make-dirs)
  ;; 加载分组映射（若哈希表为空）
  (when (zerop (hash-table-count score-sync-group-map))
    (score-sync-load-all-groups))
  ;; 获取成绩列索引
  (let* ((score-col-index (score-sync-find-score-col-index))
         ;; 提取当前行学号和成绩
         (current-no (score-sync-get-current-student-no))
         (current-score (score-sync-get-current-score score-col-index))
         ;; 获取同组学号列表
         (group-no-list (gethash current-no score-sync-group-map))
         (sync-count 0))
    ;; 检查同组学号是否存在
    (unless group-no-list
      (error "[score-sync] 错误：学号<%s>未找到对应分组信息" current-no))
    ;; 检查当前成绩是否为空
    (if (string-blank-p current-score)
        (warn "[score-sync] 提醒：当前行无有效成绩，无需同步！")
      (progn
        (message "[score-sync] 开始同步学号<%s>的成绩：%s" current-no current-score)
        ;; 遍历同组学号，批量输入成绩
        (save-excursion
          (goto-char (point-min))
          (while (not (eobp))
            (let ((current-pos (point)))
              (let* ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                                  (line-end-position)))
                    (col-list (score-sync-split-line-by-space current-line)))
                ;; 处理内容行，跳过标题行（#开头）
                (when (and col-list (not (string-blank-p current-line)) (not (string-prefix-p "#" current-line)))
                  (let ((line-no (car col-list)))
                    ;; 判断当前行学号是否在同组列表且不是当前行本身
                    (when (and (member line-no group-no-list) (not (equal line-no current-no)))
                      ;; 列数不足时补充分隔符到成绩列索引
                      (let ((new-col-list (if (< (length col-list) (1+ score-col-index))
                                              (append col-list
                                                      (make-list (- (1+ score-col-index) (length col-list)) ""))
                                            col-list)))
                        ;; 替换成绩列内容
                        (setf (nth score-col-index new-col-list) current-score)
                        ;; 用制表符分隔并重新拼接行
                        (let ((new-line (string-join new-col-list "\t")))
                          ;; 删除原行，插入新行
                          (kill-whole-line)
                          (insert new-line "\n")
                          (setq sync-count (1+ sync-count)))))))
              ;; 移动到下一行
              (goto-char current-pos)
              (forward-line 1)))))
        (message "[score-sync] 成绩同步完成！共同步 %d 名学生" sync-count)))))
;; -------------------------- 插件初始化与快捷键绑定 --------------------------
(defun score-sync-plugin-init ()
  "插件初始化函数
  1. 创建目录
  2. 加载分组映射表
  3. 绑定快捷键
  4. 添加文件打开钩子
  添加到Emacs启动钩子，自动初始化"
  (message "[score-sync] 根目录：%s" score-sync-root-dir)
  (message "[score-sync] 分组目录：%s" score-sync-groups-dir)
  (message "[score-sync] 评分标准目录：%s" score-sync-rubrics-dir)
  (score-sync-make-dirs)
  ;; 无论哈希表是否为空，都加载分组映射表
  (score-sync-load-all-groups)
  ;; 绑定全局快捷键
  (define-key global-map (kbd "C-c s g") 'score-sync-open-groups-dir)
  (define-key global-map (kbd "C-c s r") 'score-sync-open-rubrics-dir)
  (define-key global-map (kbd "C-c s a") 'score-sync-auto-fill-scores)
  (define-key global-map (kbd "C-c s s") 'score-sync-sync-score)
  ;; 添加文件打开钩子
  (add-hook 'find-file-hook 'score-sync-auto-fill-scores-when-open)
  (message "[score-sync] 插件初始化完成，快捷键：C-c s s 执行成绩同步，C-c s a 手动执行成绩预输入，C-c s g 打开分组文件夹，C-c s r 打开评分标准文件夹。"))

(defun score-sync-auto-fill-scores-when-open ()
  "当打开文件时自动执行成绩预输入
  检查文件内容是否为成绩文件格式（包含#开头的列名行）"
  (when (and (buffer-file-name) (not (string-blank-p (buffer-string))))
    ;; 确保目录存在
    (score-sync-make-dirs)
    (save-excursion
      (goto-char (point-min))
      (let ((has-header nil))
        (while (and (not has-header) (<= (point) (+ (point-min) 1000)) (not (eobp)))
          (let ((current-line (buffer-substring-no-properties (line-beginning-position)
                                                              (line-end-position))))
            (when (and (string-prefix-p "#" current-line)
                       (string-match-p "student\.no" current-line)
                       (string-match-p "student\.name" current-line)
                       (string-match-p "score" current-line))
              (setq has-header t))
            (forward-line 1)))
        (when has-header
          (score-sync-auto-fill-scores))))))

;; Emacs启动时自动初始化插件
(add-hook 'emacs-startup-hook 'score-sync-plugin-init)

;; 打开分组文件夹
(defun score-sync-open-groups-dir ()
  "打开分组文件存放目录，方便快速添加分组文件"
  (interactive)
  ;;(score-sync-make-dirs)  ; 确保目录存在
  (if (eq system-type 'windows-nt)
      (progn
        (w32-shell-execute "open" score-sync-groups-dir))
    (progn
      (dired score-sync-groups-dir))))

;; 打开评分标准文件夹
(defun score-sync-open-rubrics-dir ()
  "打开评分标准文件存放目录，方便快速添加评分标准文件"
  (interactive)
  ;;(score-sync-make-dirs)  ; 确保目录存在
  (if (eq system-type 'windows-nt)
      (progn
        (w32-shell-execute "open" score-sync-rubrics-dir))
    (progn
      (dired score-sync-rubrics-dir))))

;; 插件卸载
(defun score-sync-plugin-unload ()
  "插件卸载函数，清空映射、解绑快捷键"
  (interactive)
  (clrhash score-sync-group-map)
  (global-unset-key (kbd "C-c s s"))
  (global-unset-key (kbd "C-c s g"))
  (remove-hook 'emacs-startup-hook 'score-sync-plugin-init)
  (message "[score-sync] 插件已卸载"))

(provide 'score-sync-plugin)
