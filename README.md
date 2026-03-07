文件夹score-sync是Emacs端成绩批量输入插件，将其置于~/.emacs.d/目录下，并通过添加以下配置项以启用该插件：

;; 加载成绩同步插件
(add-to-list 'load-path "~/.emacs.d/score-sync/")
(require 'score-sync-plugin)

文件夹score_sync_plugin是Web端成绩批量输入插件，浏览器打开index.html以运行插件
