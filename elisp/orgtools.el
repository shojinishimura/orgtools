(require 'org)

(add-hook 'org-mode-hook 'orgtools-keymap)
(defun orgtools-keymap ()
  (define-key org-mode-map "\C-xp" 'orgtools-ppt-generate))

(defconst orgtools-ppt-generator "/home/shoji/workspace/orgtools/lib/ppt_generator.rb")
(defconst orgtools-keep-output-buffer t)

(defun orgtools-entry-copy ()
  (save-excursion
    (let ((start-point nil)
	  (end-point nil))
      (condition-case err
	  (outline-up-heading 1)
	(error nil))
      (setq start-point (point))
      (setq end-point (or (outline-get-next-sibling) (goto-char (point-max))))
      (buffer-substring start-point end-point))))

(defun orgtools-ppt-generate ()
  (interactive)
  (let* ((input (orgtools-entry-copy))
	 (buffer (get-buffer-create "*orgtools-ppt-generate-output*")))
    (set-buffer buffer)
    (erase-buffer)
    (let ((proc (start-process "orgtools-ppt-generate" buffer
			       "ruby" "-Ks"
			       orgtools-ppt-generator)))
      (set-process-sentinel proc 'orgtools-ppt-generate-sentinel)
      (process-send-string proc input)
      (process-send-eof))))

(defun orgtools-ppt-generate-sentinel (proc state)
  (let ((buffer (process-buffer proc))
	(status (process-status proc)))
    (cond
     ((eq status 'exit)
      (unless orgtools-keep-output-buffer
	(kill-buffer buffer)))
     (t nil))))

(provide 'orgtools)
