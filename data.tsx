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

const STORAGE_KEY = "simData";

/* =========================
   Utilities
========================= */
const toEpoch = (value: string): Epoch =>
  value ? new Date(value).getTime() : "";

const toDateTimeLocal = (value: Epoch): string =>
  value ? new Date(value).toISOString().slice(0, 16) : "";

const formatDate = (value: Epoch | number): string =>
  value
    ? new Date(value).toLocaleString(undefined, {
        year: "numeric",
        month: "short",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
      })
    : "";

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
     Load / Save
========================= */
  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) setAllRows(JSON.parse(saved));
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(allRows));
  }, [allRows]);

  /* =========================
     Filtering
========================= */
  useEffect(() => {
    if (!company) setRows(allRows);
    else {
      setRows(
        allRows.filter((row) =>
          row.businessUnit.toLowerCase().includes(company.toLowerCase())
        )
      );
    }
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
  const addRow = (): void => {
    setAllRows((prev) => [
      ...prev,
      {
        id: prev.length + 1,
        ...EMPTY_ROW,
        businessUnit: company ?? "",
        currentDate: Date.now(),
      },
    ]);
    setEditingIndex(rows.length);
  };

  const deleteRow = (index: number): void => {
    const id = rows[index].id;
    setAllRows((prev) => prev.filter((row) => row.id !== id));
    setEditingIndex(null);
  };

  const updateCell = (
    index: number,
    field: keyof SimRow,
    value: string
  ): void => {
    const id = rows[index].id;

    setAllRows((prev) =>
      prev.map((row) =>
        row.id === id
          ? {
              ...row,
              [field]:
                field === "activationDate" || field === "creationDate"
                  ? toEpoch(value)
                  : value,
            }
          : row
      )
    );
  };

  /* =========================
     File Upload (Excel)
========================= */
  const handleFileUpload = async (
    e: ChangeEvent<HTMLInputElement>
  ): Promise<void> => {
    const file = e.target.files?.[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json<Record<string, any>>(sheet);

    setAllRows((prev) => [
      ...prev,
      ...data.map((item, index) => ({
        id: prev.length + index + 1,
        subscriptionId: item["Subscription Id"] ?? "",
        msisdn: item["MSISDN"] ?? "",
        iccid: item["ICCID"] ?? "",
        imsi: item["IMSI"] ?? "",
        activationDate: toEpoch(item["Activation Date"]),
        creationDate: toEpoch(item["Subscriber Creation Date"]),
        planName: item["Subscriber Plan Name"] ?? "",
        productType: item["Product Type"] ?? "",
        businessUnit: company ?? item["Business Unit Name"] ?? "",
        status: item["Product Status"] ?? "",
        currentDate: Date.now(),
      })),
    ]);
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
        {/* Headers */}
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
                <div className="td" key={field}>
                  {editingIndex === index ? (
                    field === "activationDate" ||
                    field === "creationDate" ? (
                      <input
                        type="datetime-local"
                        value={toDateTimeLocal(row[field])}
                        onChange={(e) =>
                          updateCell(
                            index,
                            field as keyof SimRow,
                            e.target.value
                          )
                        }
                      />
                    ) : (
                      <input
                        type="text"
                        value={row[field]}
                        onChange={(e) =>
                          updateCell(
                            index,
                            field as keyof SimRow,
                            e.target.value
                          )
                        }
                      />
                    )
                  ) : field === "activationDate" ||
                    field === "creationDate" ? (
                    formatDate(row[field])
                  ) : (
                    row[field]
                  )}
                </div>
              )
            )}

            <div className="td">{formatDate(row.currentDate)}</div>

            <div className="td">
              <select
                value=""
                onChange={(e) => {
                  if (e.target.value === "edit") setEditingIndex(index);
                  if (e.target.value === "save") setEditingIndex(null);
                  if (e.target.value === "delete") deleteRow(index);
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
