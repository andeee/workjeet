(ns workjeet
  (:require [dk.ative.docjure.spreadsheet :refer :all]))

(defn column-range [start-kw end-kw]
  (let [kw-to-int (comp int first seq name)
        start (kw-to-int start-kw)
        end (kw-to-int end-kw)]
    (map (comp keyword str char) (range start (inc end)))))

(defn read-timesheet [workbook-file]
  (let [workbook (load-workbook workbook-file)
        sheet (first (sheet-seq workbook))
        header-row (first (row-seq sheet))
        columns (apply array-map (interleave (column-range :A :Z) (map read-cell header-row)))]
    (rest (select-columns columns sheet))))
