(ns workjeet
  (:require [dk.ative.docjure.spreadsheet :refer :all]
            [clojure.string :as str]))

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

(defn get-day-of-week [day-row]
  (.getDayOfWeek (get-day-date day-row)))

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
  (let [first-row (if (seq row-range-seq) (inc (last row-range-seq)) 1)]
    (conj row-range-seq first-row last-row)))

(defn get-week-row-ranges [last-days-by-week]
  (->> (map #(.getRowNum %) last-days-by-week)
       (reduce calc-row-ranges [])
       (map inc)
       (partition 2)))

(defn get-sum-fns [week-row-ranges]
  (map #(str "SUM" "(D" (str/join ":D" %) ")") week-row-ranges))

(defn set-formula-on-new-cell! [row formula]
  (let [new-cell (.createCell row (.getLastCellNum row))]
    (.setCellFormula new-cell formula)))

(defn set-hours-by-week-sums! [last-days-by-week]
  (let [sum-fns (get-sum-fns (get-week-row-ranges last-days-by-week))]
    (doall (map set-formula-on-new-cell! last-days-by-week sum-fns))))

(defn apply-sums-by-week [orig-file new-file]
  (let [workbook (load-workbook orig-file)
        last-days-by-week (-> (read-timesheet workbook) get-last-days-by-week)]
    (set-hours-by-week-sums! last-days-by-week)
    (save-workbook! new-file workbook)))
