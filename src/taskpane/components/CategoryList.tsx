import React, { useState, useCallback, useMemo } from "react";
import {
  Button,
  Badge,
  Text,
  Tooltip,
} from "@fluentui/react-components";
import {
  ChevronDownRegular,
  ChevronRightRegular,
  GridRegular,
} from "@fluentui/react-icons";
import { Category } from "../../types";

interface CategoryListProps {
  categories: Category[];
  selectedCategory: string | null;
  onSelectCategory: (categoryName: string | null) => void;
  assetCounts: Record<string, number>;
}

const CategoryList: React.FC<CategoryListProps> = ({
  categories,
  selectedCategory,
  onSelectCategory,
  assetCounts,
}) => {
  const [collapsed, setCollapsed] = useState(false);

  const totalCount = useMemo(
    () => Object.values(assetCounts).reduce((a, b) => a + b, 0),
    [assetCounts]
  );

  const toggleCollapse = useCallback(() => {
    setCollapsed((prev) => !prev);
  }, []);

  // Sort categories by count (most items first)
  const sortedCategories = useMemo(
    () => [...categories].sort((a, b) => b.count - a.count),
    [categories]
  );

  return (
    <div className="category-section">
      {/* Section Header - clickable to collapse */}
      <button
        className="category-section-header"
        onClick={toggleCollapse}
        aria-expanded={!collapsed}
        aria-label={collapsed ? "Expand categories" : "Collapse categories"}
      >
        <span className="category-section-toggle">
          {collapsed ? <ChevronRightRegular /> : <ChevronDownRegular />}
        </span>
        <Text size={200} weight="semibold" className="category-section-title">
          Categories
        </Text>
        <Badge appearance="tint" size="small" color="informative">
          {categories.length}
        </Badge>
      </button>

      {/* Collapsible category list */}
      {!collapsed && (
        <div className="category-list">
          {/* All assets button */}
          <Tooltip content={`Show all ${totalCount} assets`} relationship="label">
            <Button
              className={`category-item ${selectedCategory === null ? "category-item--active" : ""}`}
              appearance={selectedCategory === null ? "primary" : "subtle"}
              size="small"
              onClick={() => onSelectCategory(null)}
              icon={<GridRegular />}
            >
              <span className="category-item-content">
                <Text
                  size={200}
                  weight={selectedCategory === null ? "semibold" : "regular"}
                >
                  All
                </Text>
                <Badge
                  appearance="filled"
                  color={selectedCategory === null ? "brand" : "informative"}
                  size="small"
                >
                  {totalCount}
                </Badge>
              </span>
            </Button>
          </Tooltip>

          {/* Category buttons sorted by count */}
          {sortedCategories.map((category) => {
            const isActive = selectedCategory === category.name;
            return (
              <Tooltip
                key={category.id}
                content={`${category.name} — ${category.count} asset${category.count !== 1 ? "s" : ""}`}
                relationship="label"
              >
                <Button
                  className={`category-item ${isActive ? "category-item--active" : ""}`}
                  appearance={isActive ? "primary" : "subtle"}
                  size="small"
                  onClick={() =>
                    onSelectCategory(isActive ? null : category.name)
                  }
                >
                  <span className="category-item-content">
                    {category.icon && <span className="category-icon">{category.icon}</span>}
                    <Text
                      size={200}
                      weight={isActive ? "semibold" : "regular"}
                      className="category-name"
                    >
                      {category.name}
                    </Text>
                    <Badge
                      appearance="filled"
                      color={isActive ? "brand" : "informative"}
                      size="small"
                    >
                      {category.count}
                    </Badge>
                  </span>
                </Button>
              </Tooltip>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default CategoryList;
