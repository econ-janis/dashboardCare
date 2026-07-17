import * as React from "react"

export const Select = ({ children, value, onValueChange }: any) => {
  return (
    <div className="relative w-full">
      {React.Children.map(children, (child) => 
        React.isValidElement(child) 
          ? React.cloneElement(child as React.ReactElement<any>, { value, onValueChange })
          : child
      )}
    </div>
  );
};

export const SelectTrigger = ({ children, className }: any) => (
  <div className={`flex h-10 w-full items-center justify-between rounded-md border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm text-slate-700 ${className}`}>
    {children}
    {/* Icono de flecha pequeña para que parezca un select real */}
    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-slate-400"><path d="m6 9 6 6 6-6"/></svg>
  </div>
);

export const SelectValue = ({ placeholder, value }: any) => (
  <span className="truncate">{value || placeholder}</span>
);

export const SelectContent = ({ children, onValueChange }: any) => (
  <select 
    className="absolute opacity-0 inset-0 w-full h-full cursor-pointer" 
    onChange={(e) => onValueChange && onValueChange(e.target.value)}
  >
    {children}
  </select>
);

export const SelectItem = ({ value, children }: any) => (
  <option value={value}>{children}</option>
);

export const MultiSelect = ({
  options,
  selected,
  onChange,
  placeholder = "Todas",
}: {
  options: string[];
  selected: string[];
  onChange: (values: string[]) => void;
  placeholder?: string;
}) => {
  const [open, setOpen] = React.useState(false);
  const containerRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleOption = (option: string) => {
    onChange(
      selected.includes(option) ? selected.filter((v) => v !== option) : [...selected, option]
    );
  };

  const label = selected.length === 0 ? placeholder : selected.join(", ");

  return (
    <div className="relative w-full" ref={containerRef}>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className="flex h-10 w-full items-center justify-between rounded-md border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm text-slate-700"
      >
        <span className="truncate text-left" title={label}>
          {label}
        </span>
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="14"
          height="14"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          className="shrink-0 text-slate-400"
        >
          <path d="m6 9 6 6 6-6" />
        </svg>
      </button>
      {open && (
        <div className="absolute z-20 mt-1 max-h-64 w-full overflow-auto rounded-md border border-slate-200 bg-white shadow-lg">
          {options.length === 0 ? (
            <div className="px-3 py-2 text-sm text-slate-400">Sin opciones</div>
          ) : (
            options.map((option) => (
              <label
                key={option}
                className="flex cursor-pointer items-center gap-2 px-3 py-2 text-sm text-slate-700 hover:bg-slate-50"
              >
                <input
                  type="checkbox"
                  checked={selected.includes(option)}
                  onChange={() => toggleOption(option)}
                  className="h-4 w-4 rounded border-slate-300"
                />
                <span className="truncate">{option}</span>
              </label>
            ))
          )}
        </div>
      )}
    </div>
  );
};