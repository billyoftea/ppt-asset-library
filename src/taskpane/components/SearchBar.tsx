import React, { useState, useCallback, useRef, useEffect } from "react";
import {
  Input,
  Button,
  Text,
} from "@fluentui/react-components";
import {
  SearchRegular,
  DismissRegular,
} from "@fluentui/react-icons";

interface SearchBarProps {
  onSearch: (query: string) => void;
  resultCount?: number;
  totalCount?: number;
}

const DEBOUNCE_MS = 200;

const SearchBar: React.FC<SearchBarProps> = ({ onSearch, resultCount, totalCount }) => {
  const [value, setValue] = useState("");
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Debounced search to avoid excessive re-renders on rapid typing
  const debouncedSearch = useCallback(
    (query: string) => {
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }
      debounceRef.current = setTimeout(() => {
        onSearch(query);
      }, DEBOUNCE_MS);
    },
    [onSearch]
  );

  // Clean up timeout on unmount
  useEffect(() => {
    return () => {
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }
    };
  }, []);

  const handleChange = useCallback(
    (_: React.ChangeEvent<HTMLInputElement>, data: { value: string }) => {
      setValue(data.value);
      debouncedSearch(data.value);
    },
    [debouncedSearch]
  );

  const handleClear = useCallback(() => {
    setValue("");
    onSearch("");
    // Re-focus input after clearing
    inputRef.current?.focus();
  }, [onSearch]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      // Immediately trigger search on Enter
      if (e.key === "Enter") {
        if (debounceRef.current) {
          clearTimeout(debounceRef.current);
        }
        onSearch(value);
      }
      // Clear on Escape
      if (e.key === "Escape" && value) {
        handleClear();
      }
    },
    [onSearch, value, handleClear]
  );

  return (
    <div className="search-bar">
      <Input
        ref={inputRef}
        className="search-input"
        placeholder="Search by name, category, tag…"
        value={value}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
        contentBefore={<SearchRegular />}
        contentAfter={
          value ? (
            <Button
              appearance="transparent"
              icon={<DismissRegular />}
              size="small"
              onClick={handleClear}
              aria-label="Clear search"
            />
          ) : undefined
        }
      />
      {/* Result count hint */}
      {value.trim() && resultCount !== undefined && totalCount !== undefined && (
        <div className="search-result-hint">
          <Text size={100}>
            {resultCount === 0
              ? "No results found"
              : `${resultCount} of ${totalCount} assets`}
          </Text>
        </div>
      )}
    </div>
  );
};

export default SearchBar;
