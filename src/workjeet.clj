(ns workjeet
  (:require [dk.ative.docjure.spreadsheet :refer :all]
            [clojure.string :as str]
            [clojure.java.io :as io]
            [clojure.tools.trace :refer :all])
  (:import [org.apache.poi.ss.usermodel Cell]))

(defmulti read-raw-cell #(.getCellType ^Cell %))
(defmethod read-raw-cell Cell/CELL_TYPE_BLANK     [_]     nil)
(defmethod read-raw-cell Cell/CELL_TYPE_STRING    [^Cell cell]  (.getStringCellValue cell))
(defmethod read-raw-cell Cell/CELL_TYPE_BOOLEAN   [^Cell cell]  (.getBooleanCellValue cell))
(defmethod read-raw-cell Cell/CELL_TYPE_NUMERIC   [^Cell cell]  (.getNumericCellValue cell))

(defmulti set-raw-cell! (fn [^Cell cell val] (type val)))

(defmethod set-raw-cell! String [^Cell cell val]
  (.setCellValue cell ^String val))

(defmethod set-raw-cell! Number [^Cell cell val]
  (.setCellValue cell (double val)))

(defmethod set-raw-cell! Boolean [^Cell cell val]
  (.setCellValue cell ^Boolean val))

(defmethod set-raw-cell! nil [^Cell cell val]
  (let [^String null nil]
      (.setCellValue cell null)))

(defn date-formatter [pattern]
  (org.joda.time.format.DateTimeFormat/forPattern pattern))

(defn column-range [start-kw end-kw]
  (let [kw-to-int (comp int first seq name)
        start (kw-to-int start-kw)
        end (kw-to-int end-kw)]
    (map (comp keyword str char) (range start (inc end)))))

(defn get-header-row [workbook]
  (let [sheet (first (sheet-seq workbook))]
    (take 6 (map read-cell (first (row-seq sheet))))))

(defn read-timesheet [workbook]
  (let [sheet (first (sheet-seq workbook))]
    (rest (row-seq sheet))))

(defn get-day-date [day-row]
  (org.joda.time.LocalDate/fromDateFields (.getDateCellValue (first day-row))))

(defn get-week-of-year [day-row]
  (.getWeekOfWeekyear (get-day-date day-row)))

(defn get-month-of-year [day-row]
  (.getMonthOfYear (get-day-date day-row)))

(def partition-by-week
  (partial partition-by get-week-of-year))

(def partition-by-month
  (partial partition-by get-month-of-year))

(defn get-last-days-by-week [timesheet]
  (->> (partition-by-week timesheet)
       (map last)))

(defn calc-row-ranges [row-range-seq last-row]
  (if (= 1 (count row-range-seq))
    (conj row-range-seq last-row)
    (let [first-row (inc (last row-range-seq))]
      (conj row-range-seq first-row last-row))))

(defn get-week-row-ranges [first-row last-days-by-week]
  (->> (map #(.getRowNum %) last-days-by-week)
       (reduce calc-row-ranges [(.getRowNum first-row)])
       (map inc)
       (partition 2)))

(defn get-sum-fns [week-row-ranges]
  (cons (str "SUM(I3, D" (str/join ":D" (first week-row-ranges)) ")") 
        (map #(str "SUM(D" (str/join ":D" %) ")") (rest week-row-ranges))))

(defn set-formula-on-new-cell! [row formula]
  (let [week-cell (.createCell row (int (.getLastCellNum row)))
        new-cell (.createCell row (int (.getLastCellNum row)))]
    (set-cell! week-cell (str "KW " (get-week-of-year row)))
    (doto new-cell
      (apply-date-format! "[h]:mm:ss")
      (.setCellFormula formula))))

(defn set-hours-by-week-sums! [first-row last-days-by-week]
  (let [sum-fns (get-sum-fns (get-week-row-ranges first-row last-days-by-week))]
    (doall (map set-formula-on-new-cell! last-days-by-week sum-fns))))

(defn clone-cell-style! [source-cell target-cell]
  (let [new-cell-style (-> target-cell .getSheet .getWorkbook .createCellStyle)]
    (.cloneStyleFrom new-cell-style (.getCellStyle source-cell))
    (.setCellStyle target-cell new-cell-style)))

(defn copy-row! [source-row target-sheet]
  (let [target-row (.createRow target-sheet (.getPhysicalNumberOfRows target-sheet))]
    (doseq [source-cell (cell-seq source-row)]
      (let [cell-num (if (= -1 (.getLastCellNum target-row)) 0 (.getLastCellNum target-row))
            target-cell (.createCell target-row (int cell-num))]
        (set-raw-cell! target-cell (read-raw-cell source-cell))
        (clone-cell-style! source-cell target-cell)))
    target-row))

(defn copy-rows! [source-rows target-sheet]
  (doall (for [source-row source-rows] (copy-row! source-row target-sheet))))

(defn make-header [month-and-year header-row]
  [["Mitarbeiter" "Andreas Wurzer"]
   ["Monat" month-and-year]
   header-row])

(defn create-by-month-workbook [month-and-year header-row]
  (let [workbook (create-xls-workbook month-and-year
                                      (make-header month-and-year header-row))]
    (doseq [header-row (row-seq (select-sheet month-and-year workbook))]
      (set-row-style! header-row (create-cell-style! workbook {:font {:bold true}})))
    workbook))

(defn make-workjeet! [target-folder header-row day-rows]
  (let [first-day-date (-> (first day-rows) get-day-date)
        month-and-year (.print (date-formatter "MMMM yyyy") first-day-date)
        by-month-workbook (create-by-month-workbook month-and-year header-row)
        by-month-sheet (select-sheet month-and-year by-month-workbook)
        new-day-rows (copy-rows! day-rows by-month-sheet)]
    (set-hours-by-week-sums! (first new-day-rows) (get-last-days-by-week new-day-rows))
    (doseq [column-idx (range 0 9)] (.autoSizeColumn by-month-sheet column-idx))
    (save-workbook! (.getAbsolutePath (io/file target-folder (str month-and-year ".xls"))) by-month-workbook)))

(defn make-workjeets! [orig-file target-folder]
  (let [workbook (load-workbook orig-file)
        header-row (get-header-row workbook)
        months (partition-by-month (read-timesheet workbook))]
    (doseq [month months]
      (make-workjeet! target-folder header-row month))))
