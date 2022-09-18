;;; orgtbl-edit.el --- Edit spreadsheet or text-delimited file as an Org table  -*- lexical-binding: t; -*-

;; Copyright (C) 2022 Shankar Rao

;; Author: Shankar Rao <shankar.rao@gmail.com>
;; URL: https://github.com/~shankar2k/orgtbl-edit
;; Version: 0.2
;; Keywords: org, orgtbl, xls, ods, csv, tsv

;; This file is not part of GNU Emacs.

;;; Commentary:

;; This package provides the function orgtbl-edit that opens a spreadsheet or
;; text-limited file is a special buffer where it can be edited as an Org
;; table using orgtbl-mode.  In particular, all org-table- and orgtbl- commands
;; will work in the buffer.  When the buffer is saved, the table is exported
;; back to the original spreadsheet or text-delimited file in its original
;; format.
;;
;; See documentation on https://github.com/shankar2k/orgtbl-edit.

;;; License:

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; History:

;; Version 0.2 (2022-09-18):

;; - Fixed bug with CSV files that have quoted fields with commas

;; Version 0.1 (2021-10-27):

;; - Initial version

;;; Code:

;;;; Requirements

(require 'pcase)
(require 'subr-x)
(require 'org-table)
(require 'ox-odt)


;;;; Variables

(defvar orgtbl-edit-spreadsheet-formats
  '("xlsx" "xls" "ods")
  "List of spreadsheet formats that can be edited using ``orgtbl-edit''.")

(defvar orgtbl-edit-header-lines 1
  "Number of header lines to skip when detecting field separator in ``orgtbl-edit''.")

(defvar-local orgtbl-edit-separator '(4)
  "Field separator between the columns.
It can have the following values:

(4)     Use the comma as a field separator
(16)    Use a TAB as field separator
\" \"   Use a space as field separator")

(defvar-local orgtbl-edit-filename ""
  "Name of spreadsheet or text-delimited file to save this table to.")

(defvar-local orgtbl-edit-is-spreadsheet-file nil
  "Non-nil if current file being edited by orgtbl-edit is a spreadsheet.")

(defvar-local orgtbl-edit-temp-file ""
  "Name of temporary CSV file to use when exporting table to a spreadsheet file.")


;;;; Functions

(defun orgtbl-edit-guess-separator ()
  "Detect the field separator for text in the current buffer.

This moves forward ``orgtbl-edit-header-lines'' line from the
beginning of the buffer before detecting the field separator."
  (save-excursion
    (goto-char (point-min))
    (forward-line orgtbl-edit-header-lines)
    (cond
     ((not (re-search-forward "^[^\n\t]+$" nil t)) '(16))
     ((not (re-search-forward "^[^\n,]+$" nil t))  '(4))
     (t " "))))


(defun orgtbl-edit-save-function ()
  "Save table back to the original spreadsheet or text-delimited file.

This function is used by ``write-contents-functions''."
  (if-let* ((save-cmd  (pcase orgtbl-edit-separator
                         ('(16) "orgtbl-to-tsv")
                         ('(4)  "orgtbl-to-csv")
                         (" "  "orgtbl-to-generic")))
            (export-file (if orgtbl-edit-is-spreadsheet-file
                             orgtbl-edit-temp-file
                           orgtbl-edit-filename))
            (result (org-table-export export-file save-cmd))
            (ext orgtbl-edit-is-spreadsheet-file))
      (org-odt-convert export-file (car ext))
    result))


;;;; Commands

;;;###autoload
(defun orgtbl-edit (filename)
  "Edit a spreadsheet or text-delimited FILENAME as an Org table.

The contents of FILENAME are displayed in a special buffer as an
Org table, and all usual ``org-table-'' and ``orgtbl-'' commands
will work in the buffer.  When the buffer is saved, its contents
are exported back to FILENAME in its original format."
  (interactive "fSpreadsheet or CSV file: ")
  (let ((table-buffer-name (format "*orgtbl: %s*" filename)))
    (catch 'dont-overwrite
      (unless (get-buffer table-buffer-name)
        (let ((is-spreadsheet (member (file-name-extension filename)
                                      orgtbl-edit-spreadsheet-formats)))
          (when (and is-spreadsheet
                     (not (yes-or-no-p
                           (format (concat "Warning: Editing %s as an Org table"
                                           " will destroy formatting and "
                                           "formulas! Continue?")
                                   (file-name-nondirectory filename)))))
            (throw 'dont-overwrite t))
          (set-buffer (get-buffer-create table-buffer-name))
          (orgtbl-mode +1)
          (setq truncate-lines t
                orgtbl-edit-filename filename
                orgtbl-edit-is-spreadsheet-file is-spreadsheet)
          (add-hook 'write-contents-functions #'orgtbl-edit-save-function nil t)
          ;; Mark the buffer as unmodified after saving, because if this is
          ;; done in write-contents-functions, then after exporting the table,
          ;; emacs will try to save the buffer as an ordinary file.
          (add-hook 'after-save-hook #'(lambda () (set-buffer-modified-p nil))
                    nil t)
          (if is-spreadsheet
              (let ((temp-filename (concat (file-name-sans-extension filename)
                                           ".csv")))
                (org-odt-convert filename "csv")
                (unless (file-readable-p temp-filename)
                  (kill-buffer table-buffer-name)
                  (error (concat "Unable to convert %s to intermediate CSV "
                                 "format.\nPlease verify that %s%s is properly "
                                 "installed and its path is\n set correctly")
                         (file-name-nondirectory filename)
                         org-odt-convert-process
                         (if (equal org-odt-convert-process "LibreOffice")
                             " Calc" "")))
                (insert-file-contents temp-filename)
                (setq orgtbl-edit-temp-file temp-filename
                      orgtbl-edit-separator '(4)))
            (insert-file-contents filename)
            (setq orgtbl-edit-separator (orgtbl-edit-guess-separator)))
          (org-table-convert-region  (point-min) (point-max)
                                     orgtbl-edit-separator)
          (set-buffer-modified-p nil)))
      (switch-to-buffer table-buffer-name))))

;;;; Footer

(provide 'orgtbl-edit)

;;; orgtbl-edit.el ends here
