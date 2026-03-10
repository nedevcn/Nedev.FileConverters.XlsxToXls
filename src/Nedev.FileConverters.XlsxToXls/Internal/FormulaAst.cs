using System.Collections.Generic;

namespace Nedev.FileConverters.XlsxToXls.Internal
{
    // simple AST node hierarchy for formulas; used to decouple parsing from token emission
    internal abstract record AstNode;

    internal record NumberNode(double Value) : AstNode;
    internal record StringNode(string Text) : AstNode;
    internal record BoolNode(bool Value) : AstNode;
    internal record RefNode(int Row, int Col, int SheetIndex, bool HasSheet) : AstNode;
    internal record AreaNode(int Row1, int Col1, int Row2, int Col2, int SheetIndex, bool HasSheet) : AstNode;
    internal record FunctionNode(string Name, List<AstNode> Arguments) : AstNode;
    internal record OperatorNode(string Op, AstNode Left, AstNode Right) : AstNode;
    internal record UnaryOperatorNode(string Op, AstNode Operand) : AstNode;
}
