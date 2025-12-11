# -*- coding: utf-8 -*-
# Axlsx is a gem or generating excel spreadsheets with charts, images and many other features. 
# 
# acts_as_xlsx provides integration into active_record for Axlsx.
# 
require 'axlsx'

# Adding to the Axlsx module 
# @see http://github.com/randym/axlsx
module Axlsx
  # === Overview
  # This module defines the acts_as_xlsx class method and provides to_xlsx support to both AR classes and instances
  module Ar
    
    def self.included(base) # :nodoc:
      base.send :extend, ClassMethods
    end
    
    # Class methods for the mixin
    module ClassMethods

      # defines the class method to inject to_xlsx
      # @option options [Array, Symbol] columns an array of symbols defining the columns and methods to call in generating sheet data for each row.
      # @option options [String] i18n (default nil) The path to search for localization. When this is specified your i18n.t will be used to determine the labels for columns.
      # @example
      #       class MyModel < ActiveRecord::Base
      #          acts_as_xlsx :columns=> [:id, :created_at, :updated_at], :i18n => 'activerecord.attributes'
      def acts_as_xlsx(options={})
        cattr_accessor :xlsx_i18n, :xlsx_columns
        self.xlsx_i18n = options.delete(:i18n) || false
        self.xlsx_columns = options.delete(:columns)
        extend Axlsx::Ar::SingletonMethods
      end
    end

    # Singleton methods for the mixin
    module SingletonMethods

      # Maps the AR class to an Axlsx package
      # options are passed into AR find
      # @param [Array, Array] columns as an array of symbols or a symbol that defines the attributes or methods to render in the sheet.
      # @option options [Integer] header_style to apply to the first row of field names
      # @option options [Array, Symbol] types an array of Axlsx types for each cell in data rows or a single type that will be applied to all types.
      # @option options [Integer, Array] style The style to pass to Worksheet#add_row
      # @option options [String] i18n The path to i18n attributes. (usually activerecord.attributes)
      # @option options [Package] package An Axlsx::Package. When this is provided the output will be added to the package as a new sheet.  # @option options [String] name This will be used to name the worksheet added to the package. If it is not provided the name of the table name will be humanized when i18n is not specified or the I18n.t for the table name.
      # @see Worksheet#add_row
      def to_xlsx(options = {})
        if self.xlsx_columns.nil?
          self.xlsx_columns = self.column_names.map { |c| c = c.to_sym }
        end

        row_style = options.delete(:style)
        header_style = options.delete(:header_style) || row_style
        types = [options.delete(:types) || []].flatten

        i18n = if options.key?(:i18n)
                 options.delete(:i18n)
               else
                 self.xlsx_i18n
               end

        # columns => [["column_name", "header"], [...], ...]
        columns = Array.wrap(options.delete(:columns) || self.xlsx_columns).flat_map { |col|
          col.is_a?(Hash) ? col.to_a : [[col, nil]]
        }

        headers = options.delete(:headers)
        headers||= columns.map { |(col_name, col_header)|
          if col_header
            col_header
          elsif i18n == true
            self.human_attribute_name(col_name)
          elsif i18n
            I18n.t("#{i18n_key}.#{self.name.underscore}.#{col_name}", default: col_name.to_s.humanize)
          else
            col_name.to_s.humanize
          end
        }

        p = options.delete(:package) || Package.new
        row_style = p.workbook.styles.add_style(row_style) unless row_style.nil?
        header_style = p.workbook.styles.add_style(header_style) unless header_style.nil?
        i18n_key = i18n == true ? 'activerecord.attributes' : i18n
        sheet_name = options.delete(:name) || (i18n ? (i18n == true ? self.model_name.human : I18n.t("#{i18n_key}.#{table_name.underscore}", default: table_name.humanize)) : table_name.humanize)
        data = options.delete(:data) || where(options[:where]).order(options[:order]).to_a
        data = data.compact.flatten


        return p if data.empty?
        p.workbook.add_worksheet(:name=>sheet_name) do |sheet|
          sheet.add_row headers, :style=>header_style
          
          data.each do |r|
            row_data = columns.map do |(column_name, _column_header)|
              v = r
              column_name.to_s.split('.').each do |method|
                !v.nil? ? v = v.send(method) : v = nil
              end
              v
            end
            sheet.add_row row_data, :style=>row_style, :types=>types
          end
        end
        p
      end
    end
  end
end

require 'active_record'
ActiveRecord::Base.send :include, Axlsx::Ar


