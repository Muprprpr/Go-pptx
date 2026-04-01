package parts

// ============================================================================
// EmbeddingPart - 嵌入部件
// ============================================================================
//
// 对应 /ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx
// 存储嵌入的 Excel 数据（二进制文件）
//
// ============================================================================

import (
	"bytes"
	"fmt"
	"io"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
)

// EmbeddingType 嵌入类型
type EmbeddingType int

const (
	EmbeddingTypeUnknown EmbeddingType = iota
	EmbeddingTypeExcel      // Excel 工作表
	EmbeddingTypeWord       // Word 文档
	EmbeddingTypeOther      // 其他
)

// EmbeddingPart 嵌入部件
type EmbeddingPart struct {
	uri       *opc.PackURI
	data      []byte
	embedType EmbeddingType
	mu        sync.RWMutex
}

// NewEmbeddingPart 创建新的嵌入部件
func NewEmbeddingPart(id int, embedType EmbeddingType) *EmbeddingPart {
	return &EmbeddingPart{
		uri:       opc.NewPackURI(fmt.Sprintf("/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx", id)),
		embedType: embedType,
	}
}

// NewEmbeddingPartWithURI 使用指定 URI 创建嵌入部件
func NewEmbeddingPartWithURI(uri *opc.PackURI, embedType EmbeddingType) *EmbeddingPart {
	return &EmbeddingPart{
		uri:       uri,
		embedType: embedType,
	}
}

// PartURI 返回部件 URI
func (e *EmbeddingPart) PartURI() *opc.PackURI {
	return e.uri
}

// Data 返回嵌入数据
func (e *EmbeddingPart) Data() []byte {
	e.mu.RLock()
	defer e.mu.RUnlock()
	return e.data
}

// SetData 设置嵌入数据
func (e *EmbeddingPart) SetData(data []byte) {
	e.mu.Lock()
	defer e.mu.Unlock()
	e.data = data
}

// SetDataReader 从 Reader 设置数据
func (e *EmbeddingPart) SetDataReader(r io.Reader) error {
	e.mu.Lock()
	defer e.mu.Unlock()

	data, err := io.ReadAll(r)
	if err != nil {
		return fmt.Errorf("failed to read embedding data: %w", err)
	}
	e.data = data
	return nil
}

// EmbedType 返回嵌入类型
func (e *EmbeddingPart) EmbedType() EmbeddingType {
	e.mu.RLock()
	defer e.mu.RUnlock()
	return e.embedType
}

// Size 返回数据大小
func (e *EmbeddingPart) Size() int {
	e.mu.RLock()
	defer e.mu.RUnlock()
	return len(e.data)
}

// ============================================================================
// 数据写入
// ============================================================================

// WriteTo 将数据写入 Writer
func (e *EmbeddingPart) WriteTo(w io.Writer) (int64, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	if e.data == nil {
		return 0, nil
	}

	n, err := w.Write(e.data)
	return int64(n), err
}

// Reader 返回数据 Reader
func (e *EmbeddingPart) Reader() io.Reader {
	e.mu.RLock()
	defer e.mu.RUnlock()
	return bytes.NewReader(e.data)
}

// ============================================================================
// 类型检测
// ============================================================================

// DetectEmbeddingType 从文件名检测嵌入类型
func DetectEmbeddingType(filename string) EmbeddingType {
	// 简单的扩展名检测
	if len(filename) >= 5 {
		ext := filename[len(filename)-5:]
		switch ext {
		case ".xlsx":
			return EmbeddingTypeExcel
		case ".docx":
			return EmbeddingTypeWord
		}
	}
	return EmbeddingTypeUnknown
}

// IsExcel 是否为 Excel 嵌入
func (e *EmbeddingPart) IsExcel() bool {
	return e.embedType == EmbeddingTypeExcel
}

// IsWord 是否为 Word 嵌入
func (e *EmbeddingPart) IsWord() bool {
	return e.embedType == EmbeddingTypeWord
}
