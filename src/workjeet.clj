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

(defn get-week-of-year [date]
  (.getWeekOfWeekyear (org.joda.time.LocalDate/fromDateFields date)))

(defn get-duration [datetime]
  (let [end (.toDateTime (org.joda.time.LocalDateTime/fromDateFields datetime))
        start (.withTimeAtStartOfDay end)]
    (org.joda.time.Duration. start end)))

(defn sum-days [days day]
  (let [running-sum (.plus (or (-> (last days) :sum) org.joda.time.Duration/ZERO)
                           (get-duration (get day "Dauer (rel.)")))]
    (conj days (assoc day :sum running-sum))))

(defn sum-working-hours [week]
  (let [zero-excel-date (org.joda.time.LocalDateTime/fromDateFields
                         (org.apache.poi.ss.usermodel.DateUtil/getJavaDate 0.0))
        week-with-sum (reduce sum-days [] week)]
    (last (map #(assoc {}
                  :sum (.toPeriod (:sum %))
                  :kw (get-week-of-year (get % "Datum"))
                  :end-date (org.joda.time.LocalDate/fromDateFields (get % "Datum")))
               week-with-sum))))

(defn partition-by-week [timesheet]
  (partition-by
   #(get-week-of-year (get % "Datum"))
   timesheet))

(defn weekly-working-hours [timesheet]
  (map sum-working-hours (partition-by-week timesheet)))

