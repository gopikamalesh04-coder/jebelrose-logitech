import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type FormEvent,
} from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

type ContainerType = "20ft" | "40ft" | "40HC";
type CurrencyCode = "USD" | "OMR" | "CNY";
type ClearanceStatus =
  | "Planned"
  | "Booked"
  | "In Transit"
  | "Arrived Oman"
  | "Under Clearance"
  | "Cleared"
  | "Delivered";

type AppPage = "planner" | "private-label" | "categories" | "shipments";
type ProductCategory = "Tissues" | "Paper Products" | "Home & Kitchen" | "Personal Care" | "Office Supplies" | "Other";

type Product = {
  id: string;
  name: string;
  sku: string;
  supplier: string;
  description: string;
  unitCost: number;
  unitCostCurrency: CurrencyCode;
  moq: number;
  unitsPerCarton: number;
  cartonLength: number;
  cartonWidth: number;
  cartonHeight: number;
  category: ProductCategory;
  image?: string;
  imageName?: string;
  createdAt: string;
};

type ShipmentRecord = {
  id: string;
  containerNumber: string;
  containerType: ContainerType;
  port: string;
  eta: string;
  arrivalDate: string;
  clearanceDate: string;
  status: ClearanceStatus;
  clearingAgent: string;
  clearanceCost: number;
  customsDuty: number;
  shipmentCurrency: CurrencyCode;
  linkedProductIds: string[];
  remarks: string;
  createdAt: string;
};

type ProductFormState = {
  name: string;
  sku: string;
  supplier: string;
  description: string;
  unitCost: string;
  unitCostCurrency: CurrencyCode;
  moq: string;
  unitsPerCarton: string;
  cartonLength: string;
  cartonWidth: string;
  cartonHeight: string;
  category: ProductCategory;
  image?: string;
  imageName?: string;
};

type PlannerFormState = {
  selectedProductId: string;
  containerType: ContainerType;
  targetUnits: string;
  unitsPerCarton: string;
  cartonLength: string;
  cartonWidth: string;
  cartonHeight: string;
  reservePercent: string;
};

type ShipmentFormState = {
  containerNumber: string;
  containerType: ContainerType;
  port: string;
  eta: string;
  arrivalDate: string;
  clearanceDate: string;
  status: ClearanceStatus;
  clearingAgent: string;
  clearanceCost: string;
  customsDuty: string;
  shipmentCurrency: CurrencyCode;
  linkedProductIds: string[];
  remarks: string;
};

type ProductEntryMode = "manual" | "photo";

const PRODUCTS_STORAGE_KEY = "oman-purchasing-products";
const SHIPMENTS_STORAGE_KEY = "oman-purchasing-shipments";
const SETTINGS_STORAGE_KEY = "jebel-rose-settings";

type AppSettings = {
  displayCurrency: CurrencyCode;
  rates: Record<CurrencyCode, number>; // value of 1 USD in this currency
  backupEmail: string;
};

const defaultSettings: AppSettings = {
  displayCurrency: "USD",
  rates: { USD: 1, OMR: 0.385, CNY: 7.25 },
  backupEmail: "",
};

const currencyMeta: Record<CurrencyCode, { label: string; symbol: string; flag: string; locale: string }> = {
  USD: { label: "US Dollar", symbol: "$", flag: "🇺🇸", locale: "en-US" },
  OMR: { label: "Omani Rial", symbol: "ر.ع.", flag: "🇴🇲", locale: "en-OM" },
  CNY: { label: "Chinese Yuan", symbol: "¥", flag: "🇨🇳", locale: "zh-CN" },
};
const currencyOptions: CurrencyCode[] = ["USD", "OMR", "CNY"];

const containerProfiles: Record<
  ContainerType,
  { label: string; capacityCbm: number; maxPayloadKg: number; note: string }
> = {
  "20ft": {
    label: "20 ft container",
    capacityCbm: 33,
    maxPayloadKg: 28200,
    note: "Good for smaller purchase batches and faster replenishment.",
  },
  "40ft": {
    label: "40 ft container",
    capacityCbm: 67,
    maxPayloadKg: 26800,
    note: "Balanced option for larger orders.",
  },
  "40HC": {
    label: "40 ft high cube",
    capacityCbm: 76,
    maxPayloadKg: 26500,
    note: "Best for higher-volume private-label shipments.",
  },
};

const clearanceStatuses: ClearanceStatus[] = [
  "Planned",
  "Booked",
  "In Transit",
  "Arrived Oman",
  "Under Clearance",
  "Cleared",
  "Delivered",
];

const productCategories: ProductCategory[] = [
  "Tissues",
  "Paper Products",
  "Home & Kitchen",
  "Personal Care",
  "Office Supplies",
  "Other",
];

const statusStyles: Record<ClearanceStatus, string> = {
  Planned: "bg-slate-100 text-slate-700 ring-slate-200",
  Booked: "bg-blue-50 text-blue-700 ring-blue-200",
  "In Transit": "bg-indigo-50 text-indigo-700 ring-indigo-200",
  "Arrived Oman": "bg-amber-50 text-amber-700 ring-amber-200",
  "Under Clearance": "bg-orange-50 text-orange-700 ring-orange-200",
  Cleared: "bg-emerald-50 text-emerald-700 ring-emerald-200",
  Delivered: "bg-teal-50 text-teal-700 ring-teal-200",
};

const initialProductForm: ProductFormState = {
  name: "",
  sku: "",
  supplier: "",
  description: "",
  unitCost: "",
  unitCostCurrency: "USD",
  moq: "",
  unitsPerCarton: "",
  cartonLength: "",
  cartonWidth: "",
  cartonHeight: "",
  category: "Other",
  image: "",
  imageName: "",
};

const initialPlannerForm: PlannerFormState = {
  selectedProductId: "",
  containerType: "40HC",
  targetUnits: "",
  unitsPerCarton: "",
  cartonLength: "",
  cartonWidth: "",
  cartonHeight: "",
  reservePercent: "5",
};

const initialShipmentForm: ShipmentFormState = {
  containerNumber: "",
  containerType: "40HC",
  port: "Sohar",
  eta: "",
  arrivalDate: "",
  clearanceDate: "",
  status: "Planned",
  clearingAgent: "",
  clearanceCost: "",
  customsDuty: "",
  shipmentCurrency: "USD",
  linkedProductIds: [],
  remarks: "",
};

function readStorage<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

function writeStorage<T>(key: string, value: T) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(key, JSON.stringify(value));
}

function generateId() {
  return globalThis.crypto?.randomUUID?.() ?? `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function parseNumber(value: string) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCurrency(value: number, currency: CurrencyCode = "USD") {
  const meta = currencyMeta[currency];
  try {
    return new Intl.NumberFormat(meta.locale, {
      style: "currency",
      currency,
      maximumFractionDigits: currency === "OMR" ? 3 : 2,
    }).format(value);
  } catch {
    return `${meta.symbol} ${value.toFixed(currency === "OMR" ? 3 : 2)}`;
  }
}

function convertCurrency(value: number, from: CurrencyCode, to: CurrencyCode, rates: Record<CurrencyCode, number>) {
  if (from === to) return value;
  // rates are: 1 USD = rates[X]  ⇒  USD value = value / rates[from]
  const inUsd = value / (rates[from] || 1);
  return inUsd * (rates[to] || 1);
}

function formatCount(value: number) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(value);
}

function formatDate(value: string) {
  if (!value) return "—";
  if (value.includes("T")) return new Date(value).toLocaleString();
  return new Date(`${value}T00:00:00`).toLocaleDateString();
}

function daysSince(value: string) {
  if (!value) return null;
  const timestamp = new Date(`${value}T00:00:00`).getTime();
  return Math.max(0, Math.floor((Date.now() - timestamp) / (1000 * 60 * 60 * 24)));
}

async function fileToDataUrl(file: File) {
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") resolve(reader.result);
      else reject(new Error("Image conversion failed."));
    };
    reader.onerror = () => reject(reader.error ?? new Error("Image conversion failed."));
    reader.readAsDataURL(file);
  });
}

function SectionHeader({ eyebrow, title, description }: { eyebrow: string; title: string; description: string }) {
  return (
    <div className="space-y-2">
      <p className="text-xs font-semibold uppercase tracking-[0.24em] text-sky-600">{eyebrow}</p>
      <div className="space-y-1">
        <h2 className="text-2xl font-semibold tracking-tight text-slate-950">{title}</h2>
        <p className="max-w-2xl text-sm leading-6 text-slate-600">{description}</p>
      </div>
    </div>
  );
}

export default function App() {
  const [products, setProducts] = useState<Product[]>(() => {
    const stored = readStorage<Product[]>(PRODUCTS_STORAGE_KEY, []);
    // Migrate older records that did not include currency
    return stored.map((p) => ({ ...p, unitCostCurrency: (p as Product).unitCostCurrency || "USD" }));
  });
  const [shipments, setShipments] = useState<ShipmentRecord[]>(() => {
    const stored = readStorage<ShipmentRecord[]>(SHIPMENTS_STORAGE_KEY, []);
    return stored.map((s) => ({ ...s, shipmentCurrency: (s as ShipmentRecord).shipmentCurrency || "USD" }));
  });
  const [settings, setSettings] = useState<AppSettings>(() => {
    const stored = readStorage<Partial<AppSettings>>(SETTINGS_STORAGE_KEY, {});
    return {
      displayCurrency: stored.displayCurrency || defaultSettings.displayCurrency,
      rates: { ...defaultSettings.rates, ...(stored.rates || {}) },
      backupEmail: stored.backupEmail || "",
    };
  });
  const [showSettings, setShowSettings] = useState(false);
  const [productForm, setProductForm] = useState<ProductFormState>(initialProductForm);
  const [productEntryMode, setProductEntryMode] = useState<ProductEntryMode>("manual");
  const [currentPage, setCurrentPage] = useState<AppPage>("planner");
  const [plannerForm, setPlannerForm] = useState<PlannerFormState>(initialPlannerForm);
  const [shipmentForm, setShipmentForm] = useState<ShipmentFormState>(initialShipmentForm);
  const [notice, setNotice] = useState("");
  const [productSearch, setProductSearch] = useState("");
  const importInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => writeStorage(PRODUCTS_STORAGE_KEY, products), [products]);
  useEffect(() => writeStorage(SHIPMENTS_STORAGE_KEY, shipments), [shipments]);
  useEffect(() => writeStorage(SETTINGS_STORAGE_KEY, settings), [settings]);
  useEffect(() => {
    if (!notice) return undefined;
    const timer = window.setTimeout(() => setNotice(""), 3200);
    return () => window.clearTimeout(timer);
  }, [notice]);

  const selectedPlannerProduct = useMemo(
    () => products.find((product) => product.id === plannerForm.selectedProductId),
    [plannerForm.selectedProductId, products]
  );

  const productCartonLength = parseNumber(productForm.cartonLength);
  const productCartonWidth = parseNumber(productForm.cartonWidth);
  const productCartonHeight = parseNumber(productForm.cartonHeight);
  const productUnitsPerCarton = parseNumber(productForm.unitsPerCarton);
  const productCartonVolume = (productCartonLength * productCartonWidth * productCartonHeight) / 1_000_000;
  const hasProductCartonSize = productCartonLength > 0 && productCartonWidth > 0 && productCartonHeight > 0 && productUnitsPerCarton > 0;

  const totalRecordedProductCost = useMemo(
    () =>
      products.reduce(
        (sum, product) =>
          sum +
          convertCurrency(
            product.unitCost * Math.max(product.moq, 1),
            product.unitCostCurrency,
            settings.displayCurrency,
            settings.rates
          ),
        0
      ),
    [products, settings.displayCurrency, settings.rates]
  );

  const pendingClearanceCount = useMemo(
    () => shipments.filter((shipment) => shipment.status === "Arrived Oman" || shipment.status === "Under Clearance").length,
    [shipments]
  );

  const totalRecordedClearanceSpend = useMemo(
    () =>
      shipments.reduce(
        (sum, shipment) =>
          sum +
          convertCurrency(
            shipment.clearanceCost + shipment.customsDuty,
            shipment.shipmentCurrency,
            settings.displayCurrency,
            settings.rates
          ),
        0
      ),
    [shipments, settings.displayCurrency, settings.rates]
  );

  const plannerMetrics = useMemo(() => {
    const profile = containerProfiles[plannerForm.containerType];
    const unitsPerCarton = parseNumber(plannerForm.unitsPerCarton);
    const cartonLength = parseNumber(plannerForm.cartonLength);
    const cartonWidth = parseNumber(plannerForm.cartonWidth);
    const cartonHeight = parseNumber(plannerForm.cartonHeight);
    const reservePercent = parseNumber(plannerForm.reservePercent);
    const targetUnits = parseNumber(plannerForm.targetUnits);

    const usableCapacityCbm = profile.capacityCbm * (1 - reservePercent / 100);
    const cartonVolumeCbm = (cartonLength * cartonWidth * cartonHeight) / 1_000_000;
    const maxCartons = cartonVolumeCbm > 0 ? Math.floor(usableCapacityCbm / cartonVolumeCbm) : 0;
    const recommendedUnits = maxCartons * unitsPerCarton;
    const targetCartons = unitsPerCarton > 0 ? Math.ceil(targetUnits / unitsPerCarton) : 0;
    const targetVolumeCbm = targetCartons * cartonVolumeCbm;
    const fillPercent = usableCapacityCbm > 0 ? (targetVolumeCbm / usableCapacityCbm) * 100 : 0;
    const differenceUnits = targetUnits - recommendedUnits;
    const leftoverSpaceCbm = Math.max(usableCapacityCbm - targetVolumeCbm, 0);

    let summary = "Enter carton dimensions and units per carton to calculate order fit.";
    if (recommendedUnits > 0 && targetUnits === 0) {
      summary = `A ${profile.label} can hold about ${formatCount(recommendedUnits)} units based on the carton size entered.`;
    } else if (recommendedUnits > 0 && targetUnits > 0 && differenceUnits < 0) {
      summary = `Your target order is below the container plan by ${formatCount(Math.abs(differenceUnits))} units, leaving about ${leftoverSpaceCbm.toFixed(2)} CBM unused.`;
    } else if (recommendedUnits > 0 && targetUnits > 0 && differenceUnits > 0) {
      summary = `Your target order is over the estimated container fit by ${formatCount(differenceUnits)} units. Reduce quantity or move to a larger container.`;
    } else if (recommendedUnits > 0 && targetUnits > 0 && differenceUnits === 0) {
      summary = "Your target order is aligned with the estimated container capacity.";
    }

    return { profile, unitsPerCarton, cartonVolumeCbm, usableCapacityCbm, maxCartons, recommendedUnits, targetUnits, targetCartons, targetVolumeCbm, fillPercent, differenceUnits, leftoverSpaceCbm, summary };
  }, [plannerForm]);

  const plannerProgressWidth = `${Math.min(Math.max(plannerMetrics.fillPercent, 0), 100).toFixed(0)}%`;
  const plannerProgressTone = plannerMetrics.fillPercent > 100 ? "bg-rose-500" : plannerMetrics.fillPercent >= 85 ? "bg-emerald-500" : "bg-amber-500";

  function updateProductForm<K extends keyof ProductFormState>(key: K, value: ProductFormState[K]) {
    setProductForm((current) => ({ ...current, [key]: value }));
  }

  function updatePlannerForm<K extends keyof PlannerFormState>(key: K, value: PlannerFormState[K]) {
    setPlannerForm((current) => ({ ...current, [key]: value }));
  }

  function updateShipmentForm<K extends keyof ShipmentFormState>(key: K, value: ShipmentFormState[K]) {
    setShipmentForm((current) => ({ ...current, [key]: value }));
  }

  function loadProductIntoPlanner(product: Product) {
    setPlannerForm((current) => ({
      ...current,
      selectedProductId: product.id,
      unitsPerCarton: String(product.unitsPerCarton),
      cartonLength: String(product.cartonLength),
      cartonWidth: String(product.cartonWidth),
      cartonHeight: String(product.cartonHeight),
      targetUnits: current.targetUnits || String(product.moq || product.unitsPerCarton),
    }));
    setNotice(`${product.name} was loaded into the container calculator.`);
  }

  async function handleProductImageChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const image = await fileToDataUrl(file);
      setProductForm((current) => ({ ...current, image, imageName: file.name }));
      setProductEntryMode("photo");
      setNotice("Product image added. Continue by filling the product details and carton box size.");
    } catch {
      setNotice("Unable to load that image. Please try a different file.");
    } finally {
      event.target.value = "";
    }
  }

  function handleAddProduct(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const nextProduct: Product = {
      id: generateId(),
      name: productForm.name.trim(),
      sku: productForm.sku.trim(),
      supplier: productForm.supplier.trim(),
      description: productForm.description.trim(),
      unitCost: parseNumber(productForm.unitCost),
      unitCostCurrency: productForm.unitCostCurrency,
      moq: parseNumber(productForm.moq),
      unitsPerCarton: parseNumber(productForm.unitsPerCarton),
      cartonLength: parseNumber(productForm.cartonLength),
      cartonWidth: parseNumber(productForm.cartonWidth),
      cartonHeight: parseNumber(productForm.cartonHeight),
      category: productForm.category,
      image: productForm.image,
      imageName: productForm.imageName,
      createdAt: new Date().toISOString(),
    };
    setProducts((current) => [nextProduct, ...current]);
    setProductForm(initialProductForm);
    setProductEntryMode("manual");
    loadProductIntoPlanner(nextProduct);
    setNotice(`${nextProduct.name} was recorded in private-label products.`);
  }

  function handleRemoveProduct(productId: string, productName: string) {
    setProducts((current) => current.filter((product) => product.id !== productId));
    setShipments((current) => current.map((shipment) => ({ ...shipment, linkedProductIds: shipment.linkedProductIds.filter((id) => id !== productId) })));
    setPlannerForm((current) => current.selectedProductId === productId ? { ...current, selectedProductId: "" } : current);
    setNotice(`${productName} was removed from the product records.`);
  }

  function toggleLinkedProduct(productId: string) {
    setShipmentForm((current) => {
      const exists = current.linkedProductIds.includes(productId);
      return { ...current, linkedProductIds: exists ? current.linkedProductIds.filter((id) => id !== productId) : [...current.linkedProductIds, productId] };
    });
  }

  function handleAddShipment(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const arrivalOrLater = ["Arrived Oman", "Under Clearance", "Cleared", "Delivered"].includes(shipmentForm.status);
    const clearedOrLater = ["Cleared", "Delivered"].includes(shipmentForm.status);
    const today = new Date().toISOString().slice(0, 10);
    const nextShipment: ShipmentRecord = {
      id: generateId(),
      containerNumber: shipmentForm.containerNumber.trim(),
      containerType: shipmentForm.containerType,
      port: shipmentForm.port.trim(),
      eta: shipmentForm.eta,
      arrivalDate: shipmentForm.arrivalDate || (arrivalOrLater ? today : ""),
      clearanceDate: shipmentForm.clearanceDate || (clearedOrLater ? today : ""),
      status: shipmentForm.status,
      clearingAgent: shipmentForm.clearingAgent.trim(),
      clearanceCost: parseNumber(shipmentForm.clearanceCost),
      customsDuty: parseNumber(shipmentForm.customsDuty),
      shipmentCurrency: shipmentForm.shipmentCurrency,
      linkedProductIds: shipmentForm.linkedProductIds,
      remarks: shipmentForm.remarks.trim(),
      createdAt: new Date().toISOString(),
    };
    setShipments((current) => [nextShipment, ...current]);
    setShipmentForm(initialShipmentForm);
    setNotice(`Shipment ${nextShipment.containerNumber} was recorded.`);
  }

  function updateShipmentStatus(id: string, status: ClearanceStatus) {
    const today = new Date().toISOString().slice(0, 10);
    setShipments((current) =>
      current.map((shipment) => {
        if (shipment.id !== id) return shipment;
        return {
          ...shipment,
          status,
          arrivalDate: (["Arrived Oman", "Under Clearance", "Cleared", "Delivered"].includes(status) && !shipment.arrivalDate) ? today : shipment.arrivalDate,
          clearanceDate: (["Cleared", "Delivered"].includes(status) && !shipment.clearanceDate) ? today : shipment.clearanceDate,
        };
      })
    );
    setNotice(`Shipment status updated to ${status}.`);
  }

  function removeShipment(id: string, containerNumber: string) {
    setShipments((current) => current.filter((shipment) => shipment.id !== id));
    setNotice(`Shipment ${containerNumber} was removed from records.`);
  }

  function getTimestamp() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}_${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}`;
  }

  function exportToExcel() {
    if (products.length === 0 && shipments.length === 0) {
      setNotice("No data to export yet. Add products or shipments first.");
      return;
    }

    const workbook = XLSX.utils.book_new();

    const productSheet = products.map((product) => ({
      "Product Name": product.name,
      SKU: product.sku,
      Category: product.category,
      Supplier: product.supplier,
      Description: product.description,
      "Unit Cost": product.unitCost,
      Currency: product.unitCostCurrency,
      [`Unit Cost (${settings.displayCurrency})`]: Number(
        convertCurrency(product.unitCost, product.unitCostCurrency, settings.displayCurrency, settings.rates).toFixed(4)
      ),
      MOQ: product.moq,
      "Units / Carton": product.unitsPerCarton,
      "Carton Length (cm)": product.cartonLength,
      "Carton Width (cm)": product.cartonWidth,
      "Carton Height (cm)": product.cartonHeight,
      "Carton Volume (CBM)": (
        (product.cartonLength * product.cartonWidth * product.cartonHeight) /
        1_000_000
      ).toFixed(4),
      "Has Photo": product.image ? "Yes" : "No",
      "Created At": formatDate(product.createdAt),
    }));
    const productWs = XLSX.utils.json_to_sheet(productSheet);
    productWs["!cols"] = [
      { wch: 28 }, { wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 40 },
      { wch: 14 }, { wch: 10 }, { wch: 14 }, { wch: 18 }, { wch: 18 },
      { wch: 18 }, { wch: 18 }, { wch: 10 }, { wch: 22 },
    ];
    XLSX.utils.book_append_sheet(workbook, productWs, "Products");

    const shipmentSheet = shipments.map((shipment) => ({
      "Container No.": shipment.containerNumber,
      "Container Type": containerProfiles[shipment.containerType].label,
      "Oman Port": shipment.port,
      Status: shipment.status,
      ETA: formatDate(shipment.eta),
      "Arrival Date": formatDate(shipment.arrivalDate),
      "Clearance Date": formatDate(shipment.clearanceDate),
      "Clearing Agent": shipment.clearingAgent,
      "Clearance Cost": shipment.clearanceCost,
      "Customs Duty": shipment.customsDuty,
      Currency: shipment.shipmentCurrency,
      "Total in Original": shipment.clearanceCost + shipment.customsDuty,
      [`Total (${settings.displayCurrency})`]: Number(
        convertCurrency(
          shipment.clearanceCost + shipment.customsDuty,
          shipment.shipmentCurrency,
          settings.displayCurrency,
          settings.rates
        ).toFixed(4)
      ),
      "Linked Products": shipment.linkedProductIds
        .map((id) => products.find((p) => p.id === id)?.name || "Unknown")
        .join(", "),
      Remarks: shipment.remarks,
      "Created At": formatDate(shipment.createdAt),
    }));
    const shipmentWs = XLSX.utils.json_to_sheet(shipmentSheet);
    shipmentWs["!cols"] = [
      { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 14 },
      { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 18 },
      { wch: 14 }, { wch: 36 }, { wch: 36 }, { wch: 22 },
    ];
    XLSX.utils.book_append_sheet(workbook, shipmentWs, "Shipments");

    const summarySheet = [
      { Metric: "Company", Value: "Jebel Rose Trading" },
      { Metric: "Report Date", Value: new Date().toLocaleString() },
      { Metric: "Total Products", Value: products.length },
      { Metric: "Total Shipments", Value: shipments.length },
      { Metric: "Pending Clearance", Value: pendingClearanceCount },
      { Metric: "Total MOQ Cost (USD)", Value: totalRecordedProductCost.toFixed(2) },
      { Metric: "Total Clearance Spend (USD)", Value: totalRecordedClearanceSpend.toFixed(2) },
      {
        Metric: "Grand Total (USD)",
        Value: (totalRecordedProductCost + totalRecordedClearanceSpend).toFixed(2),
      },
    ];
    const summaryWs = XLSX.utils.json_to_sheet(summarySheet);
    summaryWs["!cols"] = [{ wch: 32 }, { wch: 36 }];
    XLSX.utils.book_append_sheet(workbook, summaryWs, "Summary");

    XLSX.writeFile(workbook, `JebelRose_Data_${getTimestamp()}.xlsx`);
    setNotice("Excel file downloaded successfully.");
  }

  async function exportToPDF() {
    if (products.length === 0 && shipments.length === 0) {
      setNotice("No data to export yet. Add products or shipments first.");
      return;
    }

    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();

    doc.setFillColor(15, 76, 58);
    doc.rect(0, 0, pageWidth, 90, "F");

    try {
      const logoData = await fetch("/jebel-rose-logo.png")
        .then((res) => res.blob())
        .then(
          (blob) =>
            new Promise<string>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(String(reader.result));
              reader.onerror = reject;
              reader.readAsDataURL(blob);
            })
        );
      doc.addImage(logoData, "PNG", 35, 12, 70, 66);
    } catch {
      // Continue without logo if it fails to load
    }

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(24);
    doc.setFont("helvetica", "bold");
    doc.text("Jebel Rose Trading", 120, 38);
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text("Trusted Quality, Naturally  |  جبل الورد للتجارة", 120, 56);
    doc.text("Purchasing & Shipment Control Report", 120, 72);
    doc.setFontSize(9);
    doc.text(`Generated: ${new Date().toLocaleString()}`, pageWidth - 220, 72);

    doc.setTextColor(30, 30, 30);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.text("Summary", 40, 120);

    autoTable(doc, {
      startY: 130,
      theme: "grid",
      head: [["Metric", "Value"]],
      body: [
        ["Total Products", String(products.length)],
        ["Total Shipments", String(shipments.length)],
        ["Pending Clearance", String(pendingClearanceCount)],
        [`Total MOQ Cost (${settings.displayCurrency})`, formatCurrency(totalRecordedProductCost, settings.displayCurrency)],
        [`Total Clearance Spend (${settings.displayCurrency})`, formatCurrency(totalRecordedClearanceSpend, settings.displayCurrency)],
        [`Grand Total (${settings.displayCurrency})`, formatCurrency(totalRecordedProductCost + totalRecordedClearanceSpend, settings.displayCurrency)],
      ],
      headStyles: { fillColor: [136, 19, 55], textColor: 255 },
      styles: { fontSize: 10 },
    });

    if (products.length > 0) {
      doc.addPage();
      doc.setFontSize(16);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(15, 76, 58);
      doc.text("Private Label Products", 40, 40);

      autoTable(doc, {
        startY: 60,
        theme: "striped",
        head: [["Name", "SKU", "Category", "Supplier", "Cost", "MOQ", "U/Ctn", "Carton (LxWxH cm)", "CBM"]],
        body: products.map((p) => [
          p.name,
          p.sku,
          p.category,
          p.supplier || "—",
          formatCurrency(p.unitCost, p.unitCostCurrency),
          formatCount(p.moq),
          String(p.unitsPerCarton),
          `${p.cartonLength}×${p.cartonWidth}×${p.cartonHeight}`,
          ((p.cartonLength * p.cartonWidth * p.cartonHeight) / 1_000_000).toFixed(3),
        ]),
        headStyles: { fillColor: [136, 19, 55], textColor: 255 },
        styles: { fontSize: 8, cellPadding: 4 },
        columnStyles: { 0: { cellWidth: 110 }, 3: { cellWidth: 80 } },
      });
    }

    if (shipments.length > 0) {
      doc.addPage();
      doc.setFontSize(16);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(15, 76, 58);
      doc.text("Oman Shipment Clearance Records", 40, 40);

      autoTable(doc, {
        startY: 60,
        theme: "striped",
        head: [["Container", "Type", "Port", "Status", "ETA", "Arrival", "Cleared", "Agent", "Cost", "Duty", "Total"]],
        body: shipments.map((s) => [
          s.containerNumber,
          containerProfiles[s.containerType].label,
          s.port,
          s.status,
          formatDate(s.eta),
          formatDate(s.arrivalDate),
          formatDate(s.clearanceDate),
          s.clearingAgent || "—",
          formatCurrency(s.clearanceCost, s.shipmentCurrency),
          formatCurrency(s.customsDuty, s.shipmentCurrency),
          formatCurrency(s.clearanceCost + s.customsDuty, s.shipmentCurrency),
        ]),
        headStyles: { fillColor: [136, 19, 55], textColor: 255 },
        styles: { fontSize: 8, cellPadding: 4 },
      });
    }

    const pageCount = doc.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(8);
      doc.setTextColor(120);
      doc.text(
        `Jebel Rose Trading  |  Page ${i} of ${pageCount}`,
        pageWidth / 2,
        doc.internal.pageSize.getHeight() - 20,
        { align: "center" }
      );
    }

    doc.save(`JebelRose_Report_${getTimestamp()}.pdf`);
    setNotice("PDF report downloaded successfully.");
  }

  function exportToJSON() {
    const backup = {
      exportedAt: new Date().toISOString(),
      company: "Jebel Rose Trading",
      version: 1,
      products,
      shipments,
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `JebelRose_Backup_${getTimestamp()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setNotice("Full backup file downloaded. Keep it safe.");
  }

  async function emailBackup() {
    if (products.length === 0 && shipments.length === 0) {
      setNotice("No data to email yet. Add products or shipments first.");
      return;
    }

    const recipient = settings.backupEmail.trim();
    if (!recipient) {
      setShowSettings(true);
      setNotice("Please add your backup email address in Settings first.");
      return;
    }

    const subject = `Jebel Rose Backup — ${new Date().toLocaleDateString()}`;
    const summary =
      `Jebel Rose Trading — Data Backup\n` +
      `Generated: ${new Date().toLocaleString()}\n\n` +
      `Summary\n` +
      `-----------------------------\n` +
      `Total Products: ${products.length}\n` +
      `Total Shipments: ${shipments.length}\n` +
      `Pending Clearance: ${pendingClearanceCount}\n` +
      `Total MOQ Cost: ${formatCurrency(totalRecordedProductCost, settings.displayCurrency)}\n` +
      `Total Clearance Spend: ${formatCurrency(totalRecordedClearanceSpend, settings.displayCurrency)}\n` +
      `Grand Total: ${formatCurrency(totalRecordedProductCost + totalRecordedClearanceSpend, settings.displayCurrency)}\n\n` +
      `Top 10 Products\n` +
      `-----------------------------\n` +
      products
        .slice(0, 10)
        .map(
          (p, i) =>
            `${i + 1}. ${p.name} (${p.sku}) — ${formatCurrency(p.unitCost, p.unitCostCurrency)} · ${p.category}`
        )
        .join("\n") +
      (products.length > 10 ? `\n... and ${products.length - 10} more` : "") +
      `\n\nRecent Shipments\n` +
      `-----------------------------\n` +
      shipments
        .slice(0, 10)
        .map(
          (s, i) =>
            `${i + 1}. ${s.containerNumber} · ${s.status} · ${s.port} · ${formatCurrency(s.clearanceCost + s.customsDuty, s.shipmentCurrency)}`
        )
        .join("\n") +
      (shipments.length > 10 ? `\n... and ${shipments.length - 10} more` : "") +
      `\n\n--\nNOTE: Please attach the Excel/PDF/JSON backup files (downloaded automatically) to this email before sending.\n\nJebel Rose Trading — جبل الورد للتجارة\nTrusted Quality, Naturally.`;

    // Build files
    const jsonBlob = new Blob(
      [JSON.stringify({ exportedAt: new Date().toISOString(), company: "Jebel Rose Trading", version: 1, products, shipments }, null, 2)],
      { type: "application/json" }
    );
    const jsonFile = new File([jsonBlob], `JebelRose_Backup_${getTimestamp()}.json`, { type: "application/json" });

    // Try native share with files first (works great on mobile)
    const nav = navigator as Navigator & {
      canShare?: (data?: ShareData) => boolean;
      share?: (data?: ShareData) => Promise<void>;
    };
    if (nav.canShare && nav.share && nav.canShare({ files: [jsonFile] })) {
      try {
        await nav.share({
          title: subject,
          text: summary,
          files: [jsonFile],
        });
        setNotice(`Backup shared. Send it to ${recipient}.`);
        return;
      } catch {
        // User cancelled or share failed → fallback
      }
    }

    // Fallback: download backup file + open default mail app
    const url = URL.createObjectURL(jsonBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = jsonFile.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    const mailto = `mailto:${encodeURIComponent(recipient)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(summary)}`;
    window.location.href = mailto;
    setNotice(`Backup file downloaded. Attach it to the email opened to ${recipient}.`);
  }

  function handleImportBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = String(e.target?.result || "");
        const data = JSON.parse(text);
        if (!Array.isArray(data.products) || !Array.isArray(data.shipments)) {
          throw new Error("Invalid backup file");
        }
        const confirmRestore = window.confirm(
          `This backup contains ${data.products.length} products and ${data.shipments.length} shipments.\n\nReplace ALL current data with this backup?`
        );
        if (!confirmRestore) return;
        setProducts(data.products);
        setShipments(data.shipments);
        setNotice("Backup restored successfully.");
      } catch {
        setNotice("Unable to read backup file. Please make sure it is a valid Jebel Rose backup.");
      } finally {
        if (importInputRef.current) importInputRef.current.value = "";
      }
    };
    reader.readAsText(file);
  }

  const categoriesWithProducts = useMemo(() => {
    const map = new Map<ProductCategory, Product[]>();
    productCategories.forEach((cat) => map.set(cat, []));
    products.forEach((product) => {
      const existing = map.get(product.category) || [];
      map.set(product.category, [...existing, product]);
    });
    return Array.from(map.entries()).map(([category, items]) => ({ category, items }));
  }, [products]);

  const filteredProducts = useMemo(() => {
    const query = productSearch.trim().toLowerCase();
    if (!query) return products;
    return products.filter((product) => {
      const haystack = [
        product.name,
        product.sku,
        product.supplier,
        product.description,
        product.category,
        String(product.unitCost),
        String(product.moq),
        String(product.unitsPerCarton),
        `${product.cartonLength}x${product.cartonWidth}x${product.cartonHeight}`,
        `${product.cartonLength} ${product.cartonWidth} ${product.cartonHeight}`,
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(query);
    });
  }, [products, productSearch]);

  function scrollToProduct(productId: string) {
    const element = document.getElementById(`product-${productId}`);
    if (element) {
      element.scrollIntoView({ behavior: "smooth", block: "center" });
      element.classList.add("ring-4", "ring-rose-300");
      window.setTimeout(() => element.classList.remove("ring-4", "ring-rose-300"), 2000);
    }
  }

  const selectedContainerProfile = containerProfiles[plannerForm.containerType];

  return (
    <div className="min-h-screen bg-stone-50 text-slate-900">
      <div className="relative isolate overflow-hidden">
        <div
          className="absolute inset-0 bg-cover bg-center"
          style={{ backgroundImage: "url('/oman-mountains-bg.jpg')" }}
        />
        <div className="absolute inset-0 bg-gradient-to-br from-emerald-950/85 via-stone-900/80 to-rose-950/85" />
        <div className="relative mx-auto max-w-7xl px-6 pb-10 pt-8 lg:px-8 lg:pb-14">
          <div className="flex flex-col items-center gap-6 text-center lg:flex-row lg:items-center lg:justify-between lg:text-left">
            <div className="flex flex-col items-center gap-5 lg:flex-row">
              <div className="flex h-32 w-44 items-center justify-center overflow-hidden rounded-3xl bg-white p-3 shadow-2xl shadow-emerald-950/50 ring-2 ring-rose-200/30 sm:h-36 sm:w-52">
                <img
                  src="/jebel-rose-logo.png"
                  alt="Jebel Rose Trading"
                  className="h-full w-full object-contain"
                />
              </div>
              <div className="space-y-1">
                <p className="text-xs font-semibold uppercase tracking-[0.32em] text-rose-200">
                  Jebel Rose Trading LLC
                </p>
                <h2 className="text-3xl font-semibold tracking-tight text-white sm:text-4xl">
                  جبل الورد للتجارة
                </h2>
                <p className="text-xs font-medium uppercase tracking-[0.28em] text-emerald-200">
                  Trusted Quality, Naturally
                </p>
              </div>
            </div>
            <div className="inline-flex items-center gap-2 rounded-full border border-rose-300/25 bg-rose-400/10 px-4 py-2 text-sm font-medium text-rose-100">
              Purchasing & Shipment Control Desk
            </div>
          </div>

          <div className="mt-8 flex flex-wrap gap-2">
            <button onClick={() => setCurrentPage("planner")} className={`rounded-full px-4 py-2 text-sm font-medium transition ${currentPage === "planner" ? "bg-white text-emerald-900" : "border border-white/10 bg-white/5 text-stone-200 hover:bg-white/10"}`}>🧮 Container Planner</button>
            <button onClick={() => setCurrentPage("private-label")} className={`rounded-full px-4 py-2 text-sm font-medium transition ${currentPage === "private-label" ? "bg-white text-emerald-900" : "border border-white/10 bg-white/5 text-stone-200 hover:bg-white/10"}`}>🏷️ Private Labelling</button>
            <button onClick={() => setCurrentPage("categories")} className={`rounded-full px-4 py-2 text-sm font-medium transition ${currentPage === "categories" ? "bg-white text-emerald-900" : "border border-white/10 bg-white/5 text-stone-200 hover:bg-white/10"}`}>📂 All Categories</button>
            <button onClick={() => setCurrentPage("shipments")} className={`rounded-full px-4 py-2 text-sm font-medium transition ${currentPage === "shipments" ? "bg-white text-emerald-900" : "border border-white/10 bg-white/5 text-stone-200 hover:bg-white/10"}`}>🛃 Oman Clearances</button>
          </div>

          <div className="mt-8 space-y-2">
            <h1 className="max-w-3xl text-3xl font-semibold tracking-tight text-white sm:text-4xl">
              {currentPage === "planner" && "Plan purchase quantities by container size"}
              {currentPage === "private-label" && "Private Labelling — All Products"}
              {currentPage === "categories" && "All Categories — Product Categories"}
              {currentPage === "shipments" && "Oman Shipment Clearance Records"}
            </h1>
          </div>
        </div>
      </div>

      <main className="mx-auto max-w-7xl px-6 py-8 lg:px-8 lg:py-10">
        {notice && <div className="mb-6 rounded-2xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-700">{notice}</div>}

        <div className="mb-6 rounded-3xl border border-rose-100 bg-gradient-to-br from-rose-50 via-white to-emerald-50 p-5 shadow-sm">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div className="flex items-start gap-3">
              <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-rose-100 text-xl">📥</div>
              <div>
                <h3 className="text-base font-semibold text-slate-900">Data Backup & Export</h3>
                <p className="text-sm text-slate-600">
                  Download your business data as Excel, PDF report, or JSON backup. Keep a copy safe on your phone or computer.
                </p>
              </div>
            </div>
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                onClick={exportToExcel}
                className="inline-flex items-center gap-2 rounded-full bg-emerald-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-emerald-700"
              >
                <span>📊</span> Excel (.xlsx)
              </button>
              <button
                type="button"
                onClick={exportToPDF}
                className="inline-flex items-center gap-2 rounded-full bg-rose-700 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-rose-800"
              >
                <span>📄</span> PDF Report
              </button>
              <button
                type="button"
                onClick={exportToJSON}
                className="inline-flex items-center gap-2 rounded-full bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800"
              >
                <span>💾</span> Full Backup (.json)
              </button>
              <label className="inline-flex cursor-pointer items-center gap-2 rounded-full border-2 border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50">
                <span>📂</span> Restore Backup
                <input
                  ref={importInputRef}
                  type="file"
                  accept="application/json,.json"
                  onChange={handleImportBackup}
                  className="hidden"
                />
              </label>
              <button
                type="button"
                onClick={emailBackup}
                className="inline-flex items-center gap-2 rounded-full bg-sky-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-sky-700"
              >
                <span>📧</span> Email Backup
              </button>
              <button
                type="button"
                onClick={() => setShowSettings(true)}
                className="inline-flex items-center gap-2 rounded-full border-2 border-emerald-700 bg-emerald-700 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-emerald-800"
                title="Currency & Email Settings"
              >
                <span>⚙️</span> Settings
              </button>
            </div>
          </div>
          {settings.backupEmail && (
            <p className="mt-3 text-xs text-slate-500">
              📧 Email backup target: <strong>{settings.backupEmail}</strong> · Display currency:{" "}
              <strong>{currencyMeta[settings.displayCurrency].flag} {settings.displayCurrency}</strong>
            </p>
          )}
        </div>

        {showSettings && (
          <div
            className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/70 p-4 backdrop-blur-sm"
            onClick={() => setShowSettings(false)}
          >
            <div
              className="w-full max-w-lg rounded-3xl bg-white p-6 shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="mb-5 flex items-start justify-between">
                <div>
                  <h3 className="text-xl font-semibold text-slate-900">⚙️ Settings</h3>
                  <p className="mt-1 text-sm text-slate-600">Configure currency, exchange rates, and backup email.</p>
                </div>
                <button
                  type="button"
                  onClick={() => setShowSettings(false)}
                  className="rounded-full p-2 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                  aria-label="Close"
                >
                  ✕
                </button>
              </div>

              <div className="space-y-5">
                <div>
                  <label className="block">
                    <span className="text-sm font-medium text-slate-700">📧 Backup email address</span>
                    <input
                      type="email"
                      value={settings.backupEmail}
                      onChange={(e) => setSettings((s) => ({ ...s, backupEmail: e.target.value }))}
                      placeholder="you@jebelrose.com"
                      className="mt-2 w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400"
                    />
                  </label>
                  <p className="mt-1 text-xs text-slate-500">
                    Used by the "Email Backup" button. Your data stays in your phone/computer — we don't store this email anywhere.
                  </p>
                </div>

                <div>
                  <span className="text-sm font-medium text-slate-700">💱 Display currency for totals</span>
                  <div className="mt-2 grid grid-cols-3 gap-2">
                    {currencyOptions.map((c) => (
                      <button
                        key={c}
                        type="button"
                        onClick={() => setSettings((s) => ({ ...s, displayCurrency: c }))}
                        className={`rounded-2xl border-2 px-3 py-3 text-sm font-medium transition ${
                          settings.displayCurrency === c
                            ? "border-emerald-600 bg-emerald-50 text-emerald-900"
                            : "border-slate-200 bg-white text-slate-700 hover:border-slate-300"
                        }`}
                      >
                        <div className="text-xl">{currencyMeta[c].flag}</div>
                        <div className="mt-1 font-semibold">{c}</div>
                        <div className="text-xs text-slate-500">{currencyMeta[c].label}</div>
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <span className="text-sm font-medium text-slate-700">📊 Exchange rates (1 USD =)</span>
                  <p className="mt-1 text-xs text-slate-500">
                    Update rates manually. Used to convert costs to your display currency.
                  </p>
                  <div className="mt-3 space-y-2">
                    {currencyOptions
                      .filter((c) => c !== "USD")
                      .map((c) => (
                        <label key={c} className="flex items-center gap-3">
                          <span className="w-24 text-sm font-medium text-slate-700">
                            {currencyMeta[c].flag} 1 USD =
                          </span>
                          <input
                            type="number"
                            min="0"
                            step="0.0001"
                            value={settings.rates[c]}
                            onChange={(e) =>
                              setSettings((s) => ({
                                ...s,
                                rates: { ...s.rates, [c]: parseNumber(e.target.value) || 0 },
                              }))
                            }
                            className="w-32 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm outline-none transition focus:border-sky-400"
                          />
                          <span className="text-sm font-medium text-slate-600">{c}</span>
                        </label>
                      ))}
                  </div>
                </div>

                <div className="rounded-2xl bg-emerald-50 p-3 text-xs text-emerald-800">
                  💡 Tip: Add your real exchange rate from your bank or xe.com for accurate conversions.
                </div>

                <button
                  type="button"
                  onClick={() => {
                    setShowSettings(false);
                    setNotice("Settings saved.");
                  }}
                  className="w-full rounded-full bg-emerald-700 px-6 py-3 text-sm font-semibold text-white transition hover:bg-emerald-800"
                >
                  Done
                </button>
              </div>
            </div>
          </div>
        )}

        {currentPage === "planner" && (
          <div className="space-y-6">
            <div className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm shadow-slate-200/70 lg:p-8">
              <SectionHeader eyebrow="Purchase planning" title="Container-size order calculator" description="Select a container size to see order recommendations based on carton dimensions." />
              <div className="mt-8 grid gap-8 xl:grid-cols-[1.05fr_0.95fr]">
                <div className="space-y-6">
                  <div className="grid gap-5 md:grid-cols-2">
                    <label className="space-y-2">
                      <span className="text-sm font-medium text-slate-700">Load saved product</span>
                      <select value={plannerForm.selectedProductId} onChange={(event) => { const nextId = event.target.value; const product = products.find((item) => item.id === nextId); if (!product) { updatePlannerForm("selectedProductId", ""); return; } loadProductIntoPlanner(product); }} className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm outline-none transition focus:border-sky-400 focus:bg-white">
                        <option value="">Manual entry</option>
                        {products.map((product) => (<option key={product.id} value={product.id}>{product.name} · {product.sku}</option>))}
                      </select>
                    </label>
                    <label className="space-y-2">
                      <span className="text-sm font-medium text-slate-700">Container type</span>
                      <select value={plannerForm.containerType} onChange={(event) => updatePlannerForm("containerType", event.target.value as ContainerType)} className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm outline-none transition focus:border-sky-400 focus:bg-white">
                        {(Object.entries(containerProfiles) as [ContainerType, (typeof containerProfiles)[ContainerType]][]).map(([key, container]) => (<option key={key} value={key}>{container.label}</option>))}
                      </select>
                    </label>
                    <label className="space-y-2">
                      <span className="text-sm font-medium text-slate-700">Target order units</span>
                      <input type="number" min="0" value={plannerForm.targetUnits} onChange={(event) => updatePlannerForm("targetUnits", event.target.value)} placeholder="Example: 12000" className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm outline-none transition focus:border-sky-400 focus:bg-white" />
                    </label>
                    <label className="space-y-2">
                      <span className="text-sm font-medium text-slate-700">Reserve space %</span>
                      <input type="number" min="0" max="20" value={plannerForm.reservePercent} onChange={(event) => updatePlannerForm("reservePercent", event.target.value)} placeholder="5" className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm outline-none transition focus:border-sky-400 focus:bg-white" />
                    </label>
                  </div>
                  <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
                    <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                      <div>
                        <h3 className="text-lg font-semibold text-slate-900">Carton data used for calculation</h3>
                        <p className="text-sm text-slate-600">{selectedPlannerProduct ? `Loaded from ${selectedPlannerProduct.name}. You can still adjust the values below.` : "Enter carton measurements manually if the product is not yet saved."}</p>
                      </div>
                      <div className="rounded-2xl bg-white px-4 py-3 text-sm text-slate-600 shadow-sm shadow-slate-200/70">{plannerMetrics.profile.note}</div>
                    </div>
                    <div className="mt-5 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Units per carton</span><input type="number" min="0" value={plannerForm.unitsPerCarton} onChange={(event) => updatePlannerForm("unitsPerCarton", event.target.value)} placeholder="24" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Carton length (cm)</span><input type="number" min="0" step="0.1" value={plannerForm.cartonLength} onChange={(event) => updatePlannerForm("cartonLength", event.target.value)} placeholder="60" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Carton width (cm)</span><input type="number" min="0" step="0.1" value={plannerForm.cartonWidth} onChange={(event) => updatePlannerForm("cartonWidth", event.target.value)} placeholder="40" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Carton height (cm)</span><input type="number" min="0" step="0.1" value={plannerForm.cartonHeight} onChange={(event) => updatePlannerForm("cartonHeight", event.target.value)} placeholder="35" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    </div>
                  </div>
                </div>
                <div className="space-y-5 rounded-3xl border border-slate-200 bg-slate-950 p-6 text-white shadow-xl shadow-slate-300/40">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <p className="text-sm font-medium text-slate-300">Recommended order for {selectedContainerProfile.label}</p>
                      <h3 className="mt-1 text-4xl font-semibold tracking-tight">{plannerMetrics.recommendedUnits > 0 ? `${formatCount(plannerMetrics.recommendedUnits)} units` : "—"}</h3>
                    </div>
                    <div className="rounded-2xl border border-white/10 bg-white/6 px-4 py-3 text-right text-sm text-slate-300">
                      <p>{selectedContainerProfile.label}</p>
                      <p>{selectedContainerProfile.capacityCbm} CBM capacity</p>
                      <p>{formatCount(selectedContainerProfile.maxPayloadKg)} kg max payload</p>
                    </div>
                  </div>
                  <div className="rounded-2xl border border-white/10 bg-white/6 p-4 text-sm leading-6 text-slate-200">{plannerMetrics.summary}</div>
                  <div className="space-y-3">
                    <div className="flex items-center justify-between text-sm text-slate-300"><span>Container fill against target order</span><span>{plannerMetrics.targetUnits > 0 ? `${plannerMetrics.fillPercent.toFixed(1)}%` : "Enter target units"}</span></div>
                    <div className="h-3 overflow-hidden rounded-full bg-white/10"><div className={`h-full rounded-full transition-all ${plannerProgressTone}`} style={{ width: plannerProgressWidth }} /></div>
                  </div>
                  <div className="grid gap-4 sm:grid-cols-2">
                    <div className="rounded-2xl border border-white/10 bg-white/6 p-4"><p className="text-sm text-slate-400">Carton volume</p><p className="mt-2 text-2xl font-semibold">{plannerMetrics.cartonVolumeCbm > 0 ? `${plannerMetrics.cartonVolumeCbm.toFixed(3)} CBM` : "—"}</p></div>
                    <div className="rounded-2xl border border-white/10 bg-white/6 p-4"><p className="text-sm text-slate-400">Maximum cartons</p><p className="mt-2 text-2xl font-semibold">{formatCount(plannerMetrics.maxCartons)}</p></div>
                    <div className="rounded-2xl border border-white/10 bg-white/6 p-4"><p className="text-sm text-slate-400">Target cartons required</p><p className="mt-2 text-2xl font-semibold">{formatCount(plannerMetrics.targetCartons)}</p></div>
                    <div className="rounded-2xl border border-white/10 bg-white/6 p-4"><p className="text-sm text-slate-400">Unused capacity</p><p className="mt-2 text-2xl font-semibold">{plannerMetrics.targetUnits > 0 ? `${plannerMetrics.leftoverSpaceCbm.toFixed(2)} CBM` : `${plannerMetrics.usableCapacityCbm.toFixed(2)} CBM`}</p></div>
                  </div>
                  <div className="grid gap-4 rounded-2xl border border-sky-400/20 bg-sky-400/10 p-4 text-sm text-sky-100 sm:grid-cols-3">
                    <div><p className="text-sky-200/80">Usable container space</p><p className="mt-1 text-lg font-semibold">{plannerMetrics.usableCapacityCbm.toFixed(2)} CBM</p></div>
                    <div><p className="text-sky-200/80">Target volume</p><p className="mt-1 text-lg font-semibold">{plannerMetrics.targetVolumeCbm.toFixed(2)} CBM</p></div>
                    <div><p className="text-sky-200/80">Difference vs fit</p><p className="mt-1 text-lg font-semibold">{plannerMetrics.targetUnits > 0 ? `${plannerMetrics.differenceUnits > 0 ? "+" : ""}${formatCount(plannerMetrics.differenceUnits)} units` : "—"}</p></div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {currentPage === "private-label" && (
          <div className="space-y-6">
            <div className="rounded-[2rem] border border-rose-100 bg-white p-6 shadow-sm shadow-rose-100/40 lg:p-8">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
                <SectionHeader eyebrow="🔍 Smart Product Search" title="Search across all product records" description="Type any product detail — name, SKU, supplier, category, cost, carton size — and the matching item will be found instantly." />
                <div className="rounded-2xl bg-rose-50 px-4 py-2 text-sm text-rose-700">
                  {products.length} total products • {filteredProducts.length} match
                </div>
              </div>
              <div className="relative mt-6">
                <div className="pointer-events-none absolute inset-y-0 left-5 flex items-center text-stone-400">
                  <svg className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                    <circle cx="11" cy="11" r="7" />
                    <path d="m21 21-4.3-4.3" />
                  </svg>
                </div>
                <input
                  type="search"
                  value={productSearch}
                  onChange={(event) => setProductSearch(event.target.value)}
                  placeholder="Search products… try item name, SKU, supplier, category, or carton size"
                  className="w-full rounded-full border-2 border-rose-200 bg-white py-4 pl-14 pr-12 text-base outline-none transition focus:border-rose-400 focus:shadow-lg focus:shadow-rose-100"
                />
                {productSearch && (
                  <button
                    type="button"
                    onClick={() => setProductSearch("")}
                    className="absolute inset-y-0 right-4 flex items-center text-stone-400 hover:text-stone-600"
                    aria-label="Clear search"
                  >
                    ✕
                  </button>
                )}
              </div>
              {productSearch && filteredProducts.length > 0 && (
                <div className="mt-4 rounded-2xl border border-rose-100 bg-rose-50/50 p-3">
                  <p className="mb-2 px-2 text-xs font-semibold uppercase tracking-[0.18em] text-rose-700">
                    Quick jump to product
                  </p>
                  <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                    {filteredProducts.slice(0, 6).map((product) => (
                      <button
                        key={product.id}
                        type="button"
                        onClick={() => scrollToProduct(product.id)}
                        className="flex items-center gap-3 rounded-xl border border-rose-100 bg-white p-3 text-left transition hover:border-rose-300 hover:shadow-md"
                      >
                        <div className="flex h-10 w-10 shrink-0 items-center justify-center overflow-hidden rounded-lg bg-stone-100">
                          {product.image ? (
                            <img src={product.image} alt={product.name} className="h-full w-full object-cover" />
                          ) : (
                            <span>📦</span>
                          )}
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="truncate text-sm font-medium text-slate-900">{product.name}</p>
                          <p className="truncate text-xs text-slate-500">
                            {product.sku} · {product.category}
                          </p>
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}
              {productSearch && filteredProducts.length === 0 && (
                <div className="mt-4 rounded-2xl border border-dashed border-rose-200 bg-rose-50/40 p-6 text-center">
                  <p className="text-sm text-rose-700">
                    No products match "<strong>{productSearch}</strong>". Try a different keyword.
                  </p>
                </div>
              )}
            </div>

            <div className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm shadow-slate-200/70 lg:p-8">
              <SectionHeader eyebrow="Product records" title="Private-label product capture" description="Add or capture product photos, then fill the required details such as SKU, description, carton size, and cost." />
              <div className="mt-8 grid gap-8 lg:grid-cols-[0.95fr_1.05fr]">
                <form onSubmit={handleAddProduct} className="space-y-5 rounded-3xl border border-slate-200 bg-slate-50 p-5">
                  <div className="rounded-3xl border border-slate-200 bg-white p-5">
                    <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
                      <div>
                        <p className="text-xs font-semibold uppercase tracking-[0.22em] text-slate-400">Item creation method</p>
                        <h3 className="mt-2 text-lg font-semibold text-slate-900">Create manually or create through photo</h3>
                        <p className="mt-2 max-w-xl text-sm leading-6 text-slate-600">Carton box size must be entered in both cases. Whether you create the item manually or capture it through photo, length, width, height, and units per carton are required.</p>
                      </div>
                      <div className="inline-flex rounded-full bg-slate-100 p-1">
                        <button type="button" onClick={() => setProductEntryMode("manual")} className={`rounded-full px-4 py-2 text-sm font-medium transition ${productEntryMode === "manual" ? "bg-slate-950 text-white" : "text-slate-600 hover:text-slate-900"}`}>Manual item</button>
                        <button type="button" onClick={() => setProductEntryMode("photo")} className={`rounded-full px-4 py-2 text-sm font-medium transition ${productEntryMode === "photo" ? "bg-slate-950 text-white" : "text-slate-600 hover:text-slate-900"}`}>Through photo</button>
                      </div>
                    </div>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-3">
                    <div className="rounded-2xl border border-slate-200 bg-white p-4"><p className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-400">Step 1</p><p className="mt-2 text-sm font-medium text-slate-800">{productEntryMode === "photo" ? "Capture or upload product" : "Enter item manually"}</p><p className="mt-1 text-sm text-slate-500">{productEntryMode === "photo" ? "Use camera or upload a product image, then complete the item details." : "You can create the item without a photo and still record full purchasing details."}</p></div>
                    <div className="rounded-2xl border border-slate-200 bg-white p-4"><p className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-400">Step 2</p><p className="mt-2 text-sm font-medium text-slate-800">Add carton box size</p><p className="mt-1 text-sm text-slate-500">Enter length, width, height, and units per carton for every item.</p></div>
                    <div className="rounded-2xl border border-slate-200 bg-white p-4"><p className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-400">Step 3</p><p className="mt-2 text-sm font-medium text-slate-800">Save product record</p><p className="mt-1 text-sm text-slate-500">Product becomes available for planning and shipment tracking.</p></div>
                  </div>
                  <div className="rounded-3xl border border-dashed border-slate-300 bg-white p-4">
                    <div className="flex flex-col gap-4 md:flex-row md:items-center">
                      <div className="flex h-36 w-full items-center justify-center overflow-hidden rounded-2xl bg-slate-100 md:w-40">
                        {productForm.image ? (<img src={productForm.image} alt="Product preview" className="h-full w-full object-cover" />) : (<div className="space-y-2 text-center text-slate-400"><div className="text-4xl">📸</div><p className="text-sm">{productEntryMode === "photo" ? "Waiting for photo" : "Photo optional"}</p></div>)}
                      </div>
                      <div className="flex-1 space-y-3">
                        <div className="flex flex-wrap gap-3">
                          <label className="inline-flex cursor-pointer items-center gap-2 rounded-full bg-slate-950 px-4 py-2 text-sm font-medium text-white transition hover:bg-slate-800"><span>Capture photo</span><input type="file" accept="image/*" capture="environment" onChange={handleProductImageChange} className="hidden" /></label>
                          <label className="inline-flex cursor-pointer items-center gap-2 rounded-full border border-slate-200 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50"><span>Upload image</span><input type="file" accept="image/*" onChange={handleProductImageChange} className="hidden" /></label>
                          {productForm.image && (<button type="button" onClick={() => { setProductForm((current) => ({ ...current, image: "", imageName: "" })); setProductEntryMode("manual"); setNotice("Product image removed. You can continue with manual item entry."); }} className="inline-flex items-center gap-2 rounded-full border border-rose-200 bg-rose-50 px-4 py-2 text-sm font-medium text-rose-700 transition hover:bg-rose-100">Remove photo</button>)}
                        </div>
                        <div className="rounded-2xl bg-sky-50 px-4 py-3 text-sm text-sky-700">{productForm.image ? `Photo captured: ${productForm.imageName || "Product image"}. Now enter the product details and carton box size below.` : productEntryMode === "photo" ? "Capture the product first, then fill all item details and the carton box size fields." : "Manual item creation is active. You can save the item without a photo, but carton box size is still required."}</div>
                      </div>
                    </div>
                  </div>
                  <div className="grid gap-4 md:grid-cols-2">
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Product name</span><input required value={productForm.name} onChange={(event) => updateProductForm("name", event.target.value)} placeholder="Private label tissue box" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">SKU / item code</span><input required value={productForm.sku} onChange={(event) => updateProductForm("sku", event.target.value)} placeholder="PL-001" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Supplier</span><input value={productForm.supplier} onChange={(event) => updateProductForm("supplier", event.target.value)} placeholder="Supplier name" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Unit cost</span><div className="flex gap-2"><input required type="number" min="0" step="0.01" value={productForm.unitCost} onChange={(event) => updateProductForm("unitCost", event.target.value)} placeholder="1.85" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /><select value={productForm.unitCostCurrency} onChange={(event) => updateProductForm("unitCostCurrency", event.target.value as CurrencyCode)} className="rounded-2xl border border-slate-200 bg-white px-3 py-3 text-sm outline-none transition focus:border-sky-400">{currencyOptions.map((c) => (<option key={c} value={c}>{currencyMeta[c].flag} {c}</option>))}</select></div></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Category</span><select value={productForm.category} onChange={(event) => updateProductForm("category", event.target.value as ProductCategory)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400">{productCategories.map((cat) => (<option key={cat} value={cat}>{cat}</option>))}</select></label>
                  </div>
                  <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Product description</span><textarea required value={productForm.description} onChange={(event) => updateProductForm("description", event.target.value)} rows={4} placeholder="Describe specification, packaging, private-label requirements, or finishing notes." className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                  <div className="rounded-3xl border border-sky-200 bg-sky-50/70 p-5">
                    <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                      <div><p className="text-xs font-semibold uppercase tracking-[0.22em] text-sky-700">Required carton data</p><h4 className="mt-2 text-lg font-semibold text-slate-900">Carton box size for manual and photo items</h4><p className="mt-2 max-w-2xl text-sm leading-6 text-slate-600">Add the carton box dimensions here whenever you create an item manually or through photo. These values are used in the container order calculator.</p></div>
                      <div className="rounded-2xl bg-white px-4 py-3 text-sm shadow-sm shadow-slate-200/60"><p className="font-medium text-slate-700">{hasProductCartonSize ? `${productCartonLength} × ${productCartonWidth} × ${productCartonHeight} cm` : "Waiting for carton size"}</p><p className="mt-1 text-slate-500">{hasProductCartonSize ? `${productUnitsPerCarton} units/carton · ${productCartonVolume.toFixed(3)} CBM` : "Length, width, height, and units/carton are required."}</p></div>
                    </div>
                    <div className="mt-5 grid gap-4 sm:grid-cols-2 xl:grid-cols-5">
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">MOQ units</span><input required type="number" min="1" value={productForm.moq} onChange={(event) => updateProductForm("moq", event.target.value)} placeholder="5000" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Units/carton</span><input required type="number" min="1" value={productForm.unitsPerCarton} onChange={(event) => updateProductForm("unitsPerCarton", event.target.value)} placeholder="24" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Carton length (cm)</span><input required type="number" min="0.1" step="0.1" value={productForm.cartonLength} onChange={(event) => updateProductForm("cartonLength", event.target.value)} placeholder="60" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Carton width (cm)</span><input required type="number" min="0.1" step="0.1" value={productForm.cartonWidth} onChange={(event) => updateProductForm("cartonWidth", event.target.value)} placeholder="40" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                      <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Carton height (cm)</span><input required type="number" min="0.1" step="0.1" value={productForm.cartonHeight} onChange={(event) => updateProductForm("cartonHeight", event.target.value)} placeholder="35" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    </div>
                  </div>
                  <button type="submit" className="inline-flex h-12 w-full items-center justify-center rounded-full bg-slate-950 px-6 text-sm font-semibold text-white transition hover:bg-slate-800">Save private-label product</button>
                </form>
                <div className="space-y-4">
                  <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5"><div className="flex items-start justify-between gap-3"><div><h3 className="text-lg font-semibold text-slate-900">Saved product records</h3><p className="mt-1 text-sm text-slate-600">Add or remove product records as your private-label catalog changes.</p></div><div className="rounded-2xl bg-white px-4 py-3 text-right shadow-sm shadow-slate-200/60"><p className="text-xs uppercase tracking-[0.18em] text-slate-400">Total MOQ cost</p><p className="mt-1 text-lg font-semibold text-slate-900">{formatCurrency(totalRecordedProductCost, settings.displayCurrency)}</p></div></div></div>
                  {products.length === 0 ? (<div className="rounded-3xl border border-dashed border-slate-300 bg-slate-50 px-6 py-10 text-center"><div className="text-4xl">🗂️</div><h3 className="mt-4 text-lg font-semibold text-slate-900">No products saved yet</h3><p className="mt-2 text-sm leading-6 text-slate-600">Capture a product photo and add the product details to build your private-label record library.</p></div>) : filteredProducts.length === 0 ? (<div className="rounded-3xl border border-dashed border-rose-300 bg-rose-50 px-6 py-10 text-center"><div className="text-4xl">🔍</div><h3 className="mt-4 text-lg font-semibold text-slate-900">No matches found</h3><p className="mt-2 text-sm leading-6 text-slate-600">Try a different search keyword or clear the search to see all products.</p></div>) : (<div className="grid gap-4">{filteredProducts.map((product) => { const cartonVolume = (product.cartonLength * product.cartonWidth * product.cartonHeight) / 1_000_000; return (<article key={product.id} id={`product-${product.id}`} className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-sm shadow-slate-200/60 transition"><div className="grid gap-0 md:grid-cols-[180px_1fr]"><div className="flex h-full min-h-[180px] items-center justify-center bg-slate-100">{product.image ? (<img src={product.image} alt={product.name} className="h-full w-full object-cover" />) : (<div className="space-y-2 text-center text-slate-400"><div className="text-4xl">📦</div><p className="text-sm">No image</p></div>)}</div><div className="space-y-5 p-5"><div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between"><div><div className="flex flex-wrap items-center gap-2"><h3 className="text-lg font-semibold text-slate-950">{product.name}</h3><span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-medium text-slate-600">{product.sku}</span><span className="rounded-full bg-sky-100 px-3 py-1 text-xs font-medium text-sky-700">{product.category}</span></div><p className="mt-1 text-sm text-slate-600">{product.supplier || "Supplier not entered"}</p></div><div className="flex flex-wrap gap-2"><button type="button" onClick={() => loadProductIntoPlanner(product)} className="rounded-full bg-slate-950 px-4 py-2 text-sm font-medium text-white transition hover:bg-slate-800">Use in calculator</button><button type="button" onClick={() => handleRemoveProduct(product.id, product.name)} className="rounded-full border border-rose-200 bg-rose-50 px-4 py-2 text-sm font-medium text-rose-700 transition hover:bg-rose-100">Remove</button></div></div><p className="text-sm leading-6 text-slate-600">{product.description}</p><div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">Unit cost</p><p className="mt-1 text-lg font-semibold text-slate-900">{formatCurrency(product.unitCost, product.unitCostCurrency)}</p>{product.unitCostCurrency !== settings.displayCurrency && (<p className="mt-0.5 text-xs text-slate-500">≈ {formatCurrency(convertCurrency(product.unitCost, product.unitCostCurrency, settings.displayCurrency, settings.rates), settings.displayCurrency)}</p>)}</div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">MOQ</p><p className="mt-1 text-lg font-semibold text-slate-900">{formatCount(product.moq)} units</p></div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">Carton setup</p><p className="mt-1 text-lg font-semibold text-slate-900">{product.unitsPerCarton} units / carton</p></div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">Carton volume</p><p className="mt-1 text-lg font-semibold text-slate-900">{cartonVolume.toFixed(3)} CBM</p></div></div><div className="flex flex-wrap items-center gap-3 text-sm text-slate-500"><span>Dimensions: {product.cartonLength} × {product.cartonWidth} × {product.cartonHeight} cm</span><span>•</span><span>Recorded {formatDate(product.createdAt)}</span></div></div></div></article>); })}</div>)}
                </div>
              </div>
            </div>
          </div>
        )}

        {currentPage === "categories" && (
          <div className="space-y-6">
            <div className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm shadow-slate-200/70 lg:p-8">
              <SectionHeader eyebrow="Product organization" title="All Categories" description="Browse products organized by category for better product management." />
              <div className="mt-8 grid gap-6 md:grid-cols-2 lg:grid-cols-3">
                {categoriesWithProducts.map(({ category, items }) => (
                  <div key={category} className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
                    <div className="mb-4 flex items-center justify-between">
                      <h3 className="text-lg font-semibold text-slate-900">{category}</h3>
                      <span className="rounded-full bg-white px-3 py-1 text-sm font-medium text-slate-700 shadow-sm">{items.length} products</span>
                    </div>
                    {items.length === 0 ? (<p className="text-sm text-slate-500">No products in this category yet.</p>) : (<div className="space-y-3">{items.slice(0, 5).map((product) => (<div key={product.id} className="rounded-2xl border border-slate-200 bg-white p-3"><div className="flex items-center gap-3"><div className="flex h-12 w-12 items-center justify-center rounded-xl bg-slate-100">{product.image ? (<img src={product.image} alt={product.name} className="h-full w-full rounded-xl object-cover" />) : (<span className="text-2xl">📦</span>)}</div><div className="flex-1"><p className="text-sm font-medium text-slate-900">{product.name}</p><p className="text-xs text-slate-500">{product.sku}</p></div></div></div>))}</div>)}
                    {items.length > 5 && (<p className="mt-3 text-xs text-slate-500">+{items.length - 5} more products</p>)}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {currentPage === "shipments" && (
          <div className="space-y-6">
            <div className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm shadow-slate-200/70 lg:p-8">
              <SectionHeader eyebrow="Oman clearance" title="Shipment arrival and clearance records" description="Record the container, related products, arrival date, and customs details. When a container reaches Oman, update it here." />
              <div className="mt-8 grid gap-8 lg:grid-cols-[0.96fr_1.04fr]">
                <form onSubmit={handleAddShipment} className="space-y-5 rounded-3xl border border-slate-200 bg-slate-50 p-5">
                  <div className="grid gap-4 md:grid-cols-2">
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Container number</span><input required value={shipmentForm.containerNumber} onChange={(event) => updateShipmentForm("containerNumber", event.target.value)} placeholder="MSCU1234567" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Container type</span><select value={shipmentForm.containerType} onChange={(event) => updateShipmentForm("containerType", event.target.value as ContainerType)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400">{(Object.entries(containerProfiles) as [ContainerType, (typeof containerProfiles)[ContainerType]][]).map(([key, container]) => (<option key={key} value={key}>{container.label}</option>))}</select></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Oman port</span><input required value={shipmentForm.port} onChange={(event) => updateShipmentForm("port", event.target.value)} placeholder="Sohar / Muscat / Salalah" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Status</span><select value={shipmentForm.status} onChange={(event) => updateShipmentForm("status", event.target.value as ClearanceStatus)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400">{clearanceStatuses.map((status) => (<option key={status} value={status}>{status}</option>))}</select></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">ETA</span><input type="date" value={shipmentForm.eta} onChange={(event) => updateShipmentForm("eta", event.target.value)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Arrival date in Oman</span><input type="date" value={shipmentForm.arrivalDate} onChange={(event) => updateShipmentForm("arrivalDate", event.target.value)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Clearing agent</span><input value={shipmentForm.clearingAgent} onChange={(event) => updateShipmentForm("clearingAgent", event.target.value)} placeholder="Agent or broker" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Clearance date</span><input type="date" value={shipmentForm.clearanceDate} onChange={(event) => updateShipmentForm("clearanceDate", event.target.value)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Clearance cost</span><input type="number" min="0" step="0.01" value={shipmentForm.clearanceCost} onChange={(event) => updateShipmentForm("clearanceCost", event.target.value)} placeholder="2500" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Customs duty</span><input type="number" min="0" step="0.01" value={shipmentForm.customsDuty} onChange={(event) => updateShipmentForm("customsDuty", event.target.value)} placeholder="1200" className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                    <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Currency</span><select value={shipmentForm.shipmentCurrency} onChange={(event) => updateShipmentForm("shipmentCurrency", event.target.value as CurrencyCode)} className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400">{currencyOptions.map((c) => (<option key={c} value={c}>{currencyMeta[c].flag} {c} — {currencyMeta[c].label}</option>))}</select></label>
                  </div>
                  <div className="space-y-3">
                    <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium text-slate-700">Linked products in container</span><span className="text-xs uppercase tracking-[0.18em] text-slate-400">Optional</span></div>
                    {products.length === 0 ? (<div className="rounded-2xl border border-dashed border-slate-300 bg-white px-4 py-4 text-sm text-slate-500">Add product records first if you want to attach products to this shipment.</div>) : (<div className="grid gap-2 rounded-2xl border border-slate-200 bg-white p-3">{products.map((product) => { const checked = shipmentForm.linkedProductIds.includes(product.id); return (<label key={product.id} className={`flex cursor-pointer items-center justify-between rounded-2xl border px-4 py-3 text-sm transition ${checked ? "border-sky-200 bg-sky-50 text-sky-800" : "border-slate-200 bg-white text-slate-700 hover:bg-slate-50"}`}><div><p className="font-medium">{product.name}</p><p className="text-xs text-slate-500">{product.sku}</p></div><input type="checkbox" checked={checked} onChange={() => toggleLinkedProduct(product.id)} className="h-4 w-4 rounded border-slate-300" /></label>); })}</div>)}
                  </div>
                  <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Shipment remarks</span><textarea value={shipmentForm.remarks} onChange={(event) => updateShipmentForm("remarks", event.target.value)} rows={4} placeholder="Record shipping notes, document issues, customs updates, or delivery remarks." className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-sky-400" /></label>
                  <button type="submit" className="inline-flex h-12 w-full items-center justify-center rounded-full bg-slate-950 px-6 text-sm font-semibold text-white transition hover:bg-slate-800">Save shipment clearance record</button>
                </form>
                <div className="space-y-4">
                  <div className="grid gap-4 sm:grid-cols-3"><div className="rounded-3xl border border-slate-200 bg-slate-50 p-5"><p className="text-sm text-slate-500">Recorded shipments</p><p className="mt-2 text-3xl font-semibold text-slate-950">{shipments.length}</p></div><div className="rounded-3xl border border-slate-200 bg-slate-50 p-5"><p className="text-sm text-slate-500">Pending clearance</p><p className="mt-2 text-3xl font-semibold text-slate-950">{pendingClearanceCount}</p></div><div className="rounded-3xl border border-slate-200 bg-slate-50 p-5"><p className="text-sm text-slate-500">Clearance spend ({settings.displayCurrency})</p><p className="mt-2 text-3xl font-semibold text-slate-950">{formatCurrency(totalRecordedClearanceSpend, settings.displayCurrency)}</p></div></div>
                  {shipments.length === 0 ? (<div className="rounded-3xl border border-dashed border-slate-300 bg-slate-50 px-6 py-10 text-center"><div className="text-4xl">🚢</div><h3 className="mt-4 text-lg font-semibold text-slate-900">No shipment records yet</h3><p className="mt-2 text-sm leading-6 text-slate-600">Add a container record now, then update it when the shipment reaches Oman and goes through clearance.</p></div>) : (<div className="grid gap-4">{shipments.map((shipment) => { const linkedProducts = products.filter((product) => shipment.linkedProductIds.includes(product.id)); const arrivalAge = daysSince(shipment.arrivalDate); return (<article key={shipment.id} className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm shadow-slate-200/60"><div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between"><div><div className="flex flex-wrap items-center gap-3"><h3 className="text-lg font-semibold text-slate-950">{shipment.containerNumber}</h3><span className={`inline-flex rounded-full px-3 py-1 text-xs font-semibold ring-1 ring-inset ${statusStyles[shipment.status]}`}>{shipment.status}</span></div><p className="mt-2 text-sm text-slate-600">{containerProfiles[shipment.containerType].label} · Port: {shipment.port}</p></div><div className="flex flex-wrap gap-2">{shipment.status !== "Arrived Oman" && shipment.status !== "Under Clearance" && shipment.status !== "Cleared" && shipment.status !== "Delivered" && (<button type="button" onClick={() => updateShipmentStatus(shipment.id, "Arrived Oman")} className="rounded-full bg-amber-500 px-4 py-2 text-sm font-medium text-white transition hover:bg-amber-400">Mark arrived Oman</button>)}{shipment.status !== "Under Clearance" && shipment.status !== "Cleared" && shipment.status !== "Delivered" && (<button type="button" onClick={() => updateShipmentStatus(shipment.id, "Under Clearance")} className="rounded-full bg-slate-950 px-4 py-2 text-sm font-medium text-white transition hover:bg-slate-800">Start clearance</button>)}{shipment.status !== "Cleared" && shipment.status !== "Delivered" && (<button type="button" onClick={() => updateShipmentStatus(shipment.id, "Cleared")} className="rounded-full border border-emerald-200 bg-emerald-50 px-4 py-2 text-sm font-medium text-emerald-700 transition hover:bg-emerald-100">Mark cleared</button>)}{shipment.status !== "Delivered" && (<button type="button" onClick={() => updateShipmentStatus(shipment.id, "Delivered")} className="rounded-full border border-teal-200 bg-teal-50 px-4 py-2 text-sm font-medium text-teal-700 transition hover:bg-teal-100">Mark delivered</button>)}<button type="button" onClick={() => removeShipment(shipment.id, shipment.containerNumber)} className="rounded-full border border-rose-200 bg-rose-50 px-4 py-2 text-sm font-medium text-rose-700 transition hover:bg-rose-100">Remove</button></div></div><div className="mt-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">ETA</p><p className="mt-1 text-base font-semibold text-slate-900">{formatDate(shipment.eta)}</p></div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">Arrival in Oman</p><p className="mt-1 text-base font-semibold text-slate-900">{formatDate(shipment.arrivalDate)}</p>{arrivalAge !== null && (<p className="mt-1 text-xs text-slate-500">{arrivalAge} days ago</p>)}</div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">Clearance date</p><p className="mt-1 text-base font-semibold text-slate-900">{formatDate(shipment.clearanceDate)}</p></div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm text-slate-500">Cost + duty</p><p className="mt-1 text-base font-semibold text-slate-900">{formatCurrency(shipment.clearanceCost + shipment.customsDuty, shipment.shipmentCurrency)}</p>{shipment.shipmentCurrency !== settings.displayCurrency && (<p className="mt-0.5 text-xs text-slate-500">≈ {formatCurrency(convertCurrency(shipment.clearanceCost + shipment.customsDuty, shipment.shipmentCurrency, settings.displayCurrency, settings.rates), settings.displayCurrency)}</p>)}</div></div><div className="mt-5 grid gap-4 md:grid-cols-[0.9fr_1.1fr]"><div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><p className="text-sm font-medium text-slate-700">Clearing agent</p><p className="mt-2 text-sm leading-6 text-slate-600">{shipment.clearingAgent || "No clearing agent recorded yet."}</p></div><div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><p className="text-sm font-medium text-slate-700">Linked products</p>{linkedProducts.length === 0 ? (<p className="mt-2 text-sm leading-6 text-slate-600">No products linked to this shipment.</p>) : (<div className="mt-3 flex flex-wrap gap-2">{linkedProducts.map((product) => (<span key={product.id} className="rounded-full bg-white px-3 py-2 text-xs font-medium text-slate-700 ring-1 ring-slate-200">{product.name}</span>))}</div>)}</div></div><div className="mt-5 rounded-2xl border border-slate-200 bg-slate-50 p-4"><p className="text-sm font-medium text-slate-700">Remarks</p><p className="mt-2 text-sm leading-6 text-slate-600">{shipment.remarks || "No remarks added yet."}</p></div></article>); })}</div>)}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
