(ns workjeet
  (:require [dk.ative.docjure.spreadsheet :refer :all]
            [clojure.string :as str]
            [clojure.java.io :as io]
            [clojure.tools.trace :refer :all])
  (:import [org.apache.poi.ss.usermodel Cell]))

(defmulti set-cell-no-fmt! (fn [^Cell cell val] (type val)))

(defmethod set-cell-no-fmt! String [^Cell cell val]
  (.setCellValue cell ^String val))

(defmethod set-cell-no-fmt! Number [^Cell cell val]
  (.setCellValue cell (double val)))

(defmethod set-cell-no-fmt! Boolean [^Cell cell val]
  (.setCellValue cell ^Boolean val))

(defmethod set-cell-no-fmt! java.util.Date [^Cell cell val]
  (.setCellValue cell ^java.util.Date val))

(defmethod set-cell-no-fmt! nil [^Cell cell val]
  (let [^String null nil]
      (.setCellValue cell null)))

(defn date-formatter [pattern]
  (org.joda.time.format.DateTimeFormat/forPattern pattern))

(defn column-range [start-kw end-kw]
  (let [kw-to-int (comp int first seq name)
        start (kw-to-int start-kw)
        end (kw-to-int end-kw)]
    (map (comp keyword str char) (range start (inc end)))))

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
  (let [first-row (if (seq row-range-seq) (inc (last row-range-seq)) 0)]
    (conj row-range-seq first-row last-row)))

(defn get-week-row-ranges [last-days-by-week]
  (->> (map #(.getRowNum %) last-days-by-week)
       (reduce calc-row-ranges [])
       (map inc)
       (partition 2)))

(defn get-sum-fns [week-row-ranges]
  (map #(str "SUM" "(D" (str/join ":D" %) ")") week-row-ranges))

(defn set-formula-on-new-cell! [row formula]
  (let [new-cell (.createCell row (int (.getLastCellNum row)))]
    (doto new-cell
      (apply-date-format! "[h]:mm:ss")
      (.setCellFormula formula))))

(defn set-hours-by-week-sums! [last-days-by-week]
  (let [sum-fns (get-sum-fns (get-week-row-ranges last-days-by-week))]
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
        (set-cell-no-fmt! target-cell (read-cell source-cell))
        (clone-cell-style! source-cell target-cell)))
    target-row))

(defn copy-rows! [source-rows target-sheet]
  (doall (for [source-row source-rows] (copy-row! source-row target-sheet))))

(defn make-workjeet! [parent-folder day-rows]
  (let [first-day-date (-> (first day-rows) get-day-date)
        month-and-year (.print (date-formatter "MMMM yyyy") first-day-date)
        by-month-workbook (create-xls-workbook month-and-year [])
        new-day-rows (copy-rows! day-rows (first (sheet-seq by-month-workbook)))]
    (set-hours-by-week-sums! (get-last-days-by-week new-day-rows))
    (save-workbook! (.getAbsolutePath (io/file parent-folder (str month-and-year ".xls"))) by-month-workbook)))

(defn make-workjeets! [orig-file]
  (let [workbook (load-workbook orig-file)
        months (partition-by-month (read-timesheet workbook))
        parent-folder (-> (io/file orig-file) .getParentFile)]
    (doseq [month months]
      (make-workjeet! parent-folder month))))


















































