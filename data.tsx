import { useEffect, useRef, useState, type ChangeEvent } from "react";
import { useParams } from "react-router-dom";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import "../styles/data.css";

/* =========================
   PDF Worker
========================= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

/* =========================
   Types
========================= */
type Epoch = number | "";

interface SimRow {
  id: number;
  subscriptionId: string;
  msisdn: string;
  iccid: string;
  imsi: string;
  activationDate: Epoch;
  creationDate: Epoch;
  planName: string;
  productType: string;
  businessUnit: string;
  status: string;
  currentDate: number;
}

/* =========================
   Constants
========================= */
const STORAGE_KEY = "simData";

const EMPTY_ROW: Omit<SimRow, "id" | "currentDate"> = {
  subscriptionId: "",
  msisdn: "",
  iccid: "",
  imsi: "",
  activationDate: "",
  creationDate: "",
  planName: "",
  productType: "",
  businessUnit: "",
  status: "",
};

/* =========================
   Utils
========================= */
const toEpoch = (val: string): Epoch =>
  val ? new Date(val).getTime() : "";

const toInputDate = (val: Epoch): string =>
  val ? new Date(val).toISOString().slice(0, 16) : "";

const formatDate = (val: Epoch | number): string =>
  val
    ? new Date(val).toLocaleString(undefined, {
        year: "numeric",
        month: "short",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
      })
    : "";

const isDateField = (field: string) =>
  field === "activationDate" || field === "creationDate";

/* =========================
   Component
========================= */
export default function DataPage() {
  const { company } = useParams<{ company: string }>();

  const [allRows, setAllRows] = useState<SimRow[]>([]);
  const [rows, setRows] = useState<SimRow[]>([]);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);

  const bottomRef = useRef<HTMLDivElement | null>(null);

  /* =========================
     Load from localStorage
  ========================= */
  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        setAllRows(Array.isArray(parsed) ? parsed : []);
      }
    } catch {
      setAllRows([]);
    }
  }, []);

  /* =========================
     Save to localStorage
  ========================= */
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(allRows));
  }, [allRows]);

  /* =========================
     Filtering (FIXED)
  ========================= */
  useEffect(() => {
    if (!company) {
      setRows(allRows);
      return;
    }

    const normalizedCompany = company.trim().toLowerCase();

    const filtered = allRows.filter((row) =>
      (row.businessUnit || "")
        .trim()
        .toLowerCase()
        .includes(normalizedCompany)
    );

    // Prevent empty screen
    setRows(filtered.length ? filtered : allRows);
  }, [company, allRows]);

  /* =========================
     Auto Scroll
  ========================= */
  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [rows]);

  /* =========================
     CRUD
  ========================= */
  const createRow = (): SimRow => ({
    id: Date.now(), // ✅ unique id
    ...EMPTY_ROW,
    businessUnit: company ? company.trim() : "default",
    currentDate: Date.now(),
  });

  const addRow = () => {
    setAllRows((prev) => [...prev, createRow()]);
    setEditingIndex(rows.length);
  };

  const deleteRow = (id: number) => {
    setAllRows((prev) => prev.filter((row) => row.id !== id));
    setEditingIndex(null);
  };

  const updateCell = (
    id: number,
    field: keyof SimRow,
    value: string
  ) => {
    setAllRows((prev) =>
      prev.map((row) =>
        row.id === id
          ? {
              ...row,
              [field]: isDateField(field)
                ? toEpoch(value)
                : value,
            }
          : row
      )
    );
  };

  /* =========================
     Excel Upload
  ========================= */
  const handleFileUpload = async (
    e: ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json<Record<string, any>>(sheet);

    const newRows: SimRow[] = data.map((item, index) => ({
      id: Date.now() + index,
      subscriptionId: item["Subscription Id"] ?? "",
      msisdn: item["MSISDN"] ?? "",
      iccid: item["ICCID"] ?? "",
      imsi: item["IMSI"] ?? "",
      activationDate: toEpoch(item["Activation Date"]),
      creationDate: toEpoch(item["Subscriber Creation Date"]),
      planName: item["Subscriber Plan Name"] ?? "",
      productType: item["Product Type"] ?? "",
      businessUnit:
        company || item["Business Unit Name"] || "default",
      status: item["Product Status"] ?? "",
      currentDate: Date.now(),
    }));

    setAllRows((prev) => [...prev, ...newRows]);
  };

  /* =========================
     Render Cell
  ========================= */
  const renderCell = (
    row: SimRow,
    index: number,
    field: keyof typeof EMPTY_ROW
  ) => {
    if (editingIndex === index) {
      return isDateField(field) ? (
        <input
          type="datetime-local"
          value={toInputDate(row[field])}
          onChange={(e) =>
            updateCell(row.id, field as keyof SimRow, e.target.value)
          }
        />
      ) : (
        <input
          type="text"
          value={row[field]}
          onChange={(e) =>
            updateCell(row.id, field as keyof SimRow, e.target.value)
          }
        />
      );
    }

    return isDateField(field)
      ? formatDate(row[field])
      : row[field];
  };

  /* =========================
     Render
  ========================= */
  return (
    <div className="container">
      <h2 className="page-title">SIM Data – {company}</h2>

      <div className="toolbar">
        <button className="btn add" onClick={addRow}>
          + Add Row
        </button>

        <label className="btn upload">
          Upload Excel
          <input type="file" hidden onChange={handleFileUpload} />
        </label>
      </div>

      <div className="table">
        {/* Header */}
        <div className="table-header">
          {[
            "ID",
            "Subscription ID",
            "MSISDN",
            "ICCID",
            "IMSI",
            "Activation Date",
            "Creation Date",
            "Plan Name",
            "Product Type",
            "Business Unit",
            "Status",
            "Current Date",
            "Actions",
          ].map((h) => (
            <div key={h} className="th">
              {h}
            </div>
          ))}
        </div>

        {/* Rows */}
        {rows.map((row, index) => (
          <div className="table-row card" key={row.id}>
            <div className="td">{row.id}</div>

            {(Object.keys(EMPTY_ROW) as (keyof typeof EMPTY_ROW)[]).map(
              (field) => (
                <div key={field} className="td">
                  {renderCell(row, index, field)}
                </div>
              )
            )}

            <div className="td">{formatDate(row.currentDate)}</div>

            <div className="td">
              <select
                defaultValue=""
                onChange={(e) => {
                  const action = e.target.value;
                  if (action === "edit") setEditingIndex(index);
                  if (action === "save") setEditingIndex(null);
                  if (action === "delete") deleteRow(row.id);
                }}
              >
                <option value="" disabled>
                  Actions
                </option>
                <option value="edit">Edit</option>
                <option value="save">Update</option>
                <option value="delete">Delete</option>
              </select>
            </div>
          </div>
        ))}

        <div ref={bottomRef} />
      </div>
    </div>
  );
}