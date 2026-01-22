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